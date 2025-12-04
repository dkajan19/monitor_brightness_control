import subprocess
import keyboard
import os
import tkinter as tk
import threading
import sys
import winreg
from PIL import Image, ImageDraw
import pystray
import ctypes
import math
import time
import json
from monitorcontrol import get_monitors

# --- CONFIGURATION ---
STEP = 10
THEME_COLORS = {
    'dark': {'bg': '#2d2d2d', 'fg': '#ffffff', 'bar_bg': '#404040', 'bar_fill': '#0078d4', 'icon_color': '#ffffff'},
    'light': {'bg': '#f3f3f3', 'fg': '#000000', 'bar_bg': '#d0d0d0', 'bar_fill': '#0078d4', 'icon_color': '#000000'}
}

# --- GLOBAL VARIABLES ---
current_theme = 'dark'
processing_busy = False
tray_icon = None
hide_timer = None
valid_monitors_data = []

# --- LOAD SETTINGS AND HOTKEYS FROM JSON, CREATE DEFAULT VALUES IF NOT EXISTS ---
settings_file = "settings.json"
default_settings = {
    "hotkeys": {
        "increase_brightness": {"key": "up", "modifier": "ctrl"},
        "decrease_brightness": {"key": "down", "modifier": "ctrl"}
    },
    "monitor_states": []
}

if not os.path.exists(settings_file):
    try:
        with open(settings_file, "w", encoding="utf-8") as f:
            json.dump(default_settings, f, indent=4)
        print(f"Created new settings.json with default values.")
    except Exception as e:
        print(f"Error creating default settings.json: {e}")

# --- LOAD SETTINGS ---
try:
    with open(settings_file, "r", encoding="utf-8") as f:
        settings = json.load(f)
except Exception as e:
    print(f"Error loading settings.json: {e}")
    settings = {}

hotkeys = settings.get("hotkeys", {})
for action, data in default_settings["hotkeys"].items():
    if action not in hotkeys:
        hotkeys[action] = data

monitor_states_saved = settings.get("monitor_states", [])


if sys.platform == "win32":
    try:
        user32 = ctypes.windll.user32
        dwmapi = ctypes.windll.dwmapi
        DWM_WINDOW_ATTRIBUTES = {
            'CORNER_PREFERENCE': 33,
            'CORNER_ROUND': 2,
            'DARK_MODE': 20
        }
    except ImportError:
        pass

# --- MONITOR INITIALIZATION AND NAMES (WMI) ---

def get_monitor_id_data_wmi():
    """
    Gets friendly names and serial numbers (if available) of monitors via PowerShell WMI.
    Uses ONLY UserFriendlyName to prevent brand duplication in the name.
    Returns list of dictionaries: [{'name': 'Name', 'serial': 'Serial_Number'}]
    """
    cmd = r"""
$Monitors = Get-WmiObject WmiMonitorID -Namespace root\wmi; 
$Results = @();
ForEach ($Monitor in $Monitors) 
{ 
    # Get friendly name (UserFriendlyName) and clean it
    $Name = [System.Text.Encoding]::ASCII.GetString($Monitor.UserFriendlyName).Replace("`0", "").Trim();
    $DisplayName = "";

    if ($Name) {
        $DisplayName = $Name
    } else {
        $DisplayName = "Unknown Monitor"
    }
    
    # Get serial number (VCP 0x88) and clean it
    $SerialBytes = $Monitor.SerialNumberID;
    $Serial = "";
    if ($SerialBytes) {
        $Serial = [System.Text.Encoding]::ASCII.GetString($SerialBytes).Replace("`0", "").Trim();
    }
    
    $Results += [PSCustomObject]@{
        Name = $DisplayName;
        Serial = $Serial;
    }
}
$Results | ConvertTo-Json -Compress
"""
    try:
        proc = subprocess.run(
            ["powershell", "-NoProfile", "-WindowStyle", "Hidden", "-Command", cmd],
            capture_output=True,
            text=True,
            check=True,
            creationflags=subprocess.CREATE_NO_WINDOW
        )
        json_output = proc.stdout.strip()
        if not json_output.startswith('['):
             json_output = f'[{json_output}]'
        return json.loads(json_output)
    except subprocess.CalledProcessError as e:
        print(f"Error getting WMI data: {e.stderr.strip()}")
        return []
    except FileNotFoundError:
        print("Error: PowerShell not found.")
        return []
    except json.JSONDecodeError:
        print(f"Error: Failed to decode data from PowerShell: {json_output}")
        return []

monitor_wmi_data = get_monitor_id_data_wmi()
for idx, item in enumerate(monitor_wmi_data, start=1):
    item["WMI_index"] = idx
    

def filter_supported_monitors():
    print("--------------------------------------------------")
    print("Searching for DDC monitors and attempting to match names...")
    supported_data = []
    
    unpaired_wmi_data = list(monitor_wmi_data)

    for i, monitor in enumerate(get_monitors()):
        monitor_index = i + 1
        final_name = f"Monitor {monitor_index}"  # základný názov s poradovým číslom
        model_name = None
        serial_str = ""
        capabilities = {}

        try:
            with monitor:

                # --------- SERIAL NUMBER ----------
                try:
                    serial_str = monitor.get_serial().strip()
                    print(f"  → DDC Serial: '{serial_str}'")
                except Exception:
                    print("  → DDC Serial: <not available>")

                # --------- CAPABILITIES ----------
                try:
                    capabilities = monitor.get_vcp_capabilities()

                    # Model z capabilities
                    if isinstance(capabilities, dict):
                        model_name = capabilities.get("model", None)
                        if model_name:
                            print(f"  → Model from capabilities: {model_name}")
                except Exception as e:
                    print(f"  → Capabilities not available: {e}")

                # --------- MATCH WMI ----------
                name_found = False
                if serial_str:
                    for wmi_idx, wmi_item in enumerate(unpaired_wmi_data):
                        wmi_serial = wmi_item.get('Serial', '')
                        print(f"  → WMI Serial: '{wmi_serial}' (match: {wmi_serial.upper() == serial_str.upper()})")

                        if wmi_serial.upper() == serial_str.upper():
                            final_name = f"{wmi_item['Name']} ({monitor_index})"
                            unpaired_wmi_data.pop(wmi_idx)
                            name_found = True
                            break

                # --------- NO WMI MATCH → use model name ----------
                if not name_found:
                    if model_name:
                        final_name = f"{model_name} ({monitor_index})"
                    print(f"DDC Monitor {monitor_index} paired with: {final_name}")

                supported_data.append({
                    'monitor_obj': monitor,
                    'cached_brightness': monitor.get_luminance(),
                    'friendly_name': final_name,
                    'serial': serial_str,
                    'model': model_name,
                    'capabilities': capabilities
                })

        except Exception as e:
            print(f"Monitor {monitor_index} failed: {e}")

    return supported_data


valid_monitors_data = filter_supported_monitors()
#monitor_states = [True] * len(valid_monitors_data)
# --- INITIALIZE MONITOR STATES ACCORDING TO SETTINGS.JSON ---
monitor_states = []
for i, data in enumerate(valid_monitors_data):
    if i < len(monitor_states_saved):
        monitor_states.append(monitor_states_saved[i])
    else:
        monitor_states.append(True)  # default state

if len(monitor_states_saved) < len(valid_monitors_data):
    settings["monitor_states"] = monitor_states
    try:
        with open(settings_file, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=4)
    except Exception as e:
        print(f"Error saving monitor_states to settings.json: {e}")


# --- WINDOWS THEME AND API FUNCTIONS ---
def is_system_dark_mode():
    if sys.platform != "win32":  
        return True
    try:
        key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                             r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
        return winreg.QueryValueEx(key, "AppsUseLightTheme")[0] == 0
    except Exception:
        return True

def set_window_attributes(window_handle, dark_mode=True):
    if sys.platform != "win32": return
    try:
        attrs = DWM_WINDOW_ATTRIBUTES
        hwnd = window_handle
        preference = ctypes.c_int(attrs['CORNER_ROUND'])
        dark_mode_bool = ctypes.c_int(1 if dark_mode else 0)
        dwmapi.DwmSetWindowAttribute(hwnd, attrs['CORNER_PREFERENCE'], ctypes.byref(preference), ctypes.sizeof(preference))
        dwmapi.DwmSetWindowAttribute(hwnd, attrs['DARK_MODE'], ctypes.byref(dark_mode_bool), ctypes.sizeof(dark_mode_bool))
    except Exception:
        pass

def update_theme_colors():
    global current_theme
    dark_mode = is_system_dark_mode()
    new_theme = 'dark' if dark_mode else 'light'
    colors = THEME_COLORS[new_theme]

    if new_theme != current_theme:
        current_theme = new_theme

    if root.winfo_exists():
        widgets = [(root, 'bg'), (main_frame, 'bg'), (content_frame, 'bg'), (progress_canvas, 'bg')]
        for widget, key in widgets: widget.configure(**{key: colors['bg']})
        sun_icon.configure(bg=colors['bg'], fg=colors['icon_color'])
        percent_label.configure(bg=colors['bg'], fg=colors['fg'])
        if sys.platform == "win32":
            set_window_attributes(user32.GetParent(root.winfo_id()), dark_mode=dark_mode)

    if tray_icon:
        tray_icon.icon = create_tray_icon_image(dark_mode)

# --- POLLING FOR THEME CHANGES ---
def poll_theme_changes(interval=2):
    last_theme = is_system_dark_mode()
    while True:
        current = is_system_dark_mode()
        if current != last_theme:
            last_theme = current
            update_theme_colors()
        time.sleep(interval)

# --- GUI FUNCTIONS ---
def draw_progress_bar(value):
    colors = THEME_COLORS[current_theme]
    progress_canvas.delete("all")
    progress_canvas.create_rectangle(0, 1, 200, 5, fill=colors['bar_bg'], width=0)
    progress_width = (value / 100) * 200
    if progress_width > 0:
        progress_canvas.create_rectangle(0, 1, progress_width, 5, fill=colors['bar_fill'], width=0)

def show_brightness(value):
    def update_gui():
        global hide_timer
        update_theme_colors()
        root.deiconify()
        percent_label.config(text=f"{value}%")
        draw_progress_bar(value)
        root.update_idletasks()
        if hide_timer:
            root.after_cancel(hide_timer)
        hide_timer = root.after(4000, root.withdraw)
    root.after(0, update_gui)

# --- BRIGHTNESS CONTROL LOGIC ---
def change_brightness_thread(delta):
    global processing_busy
    if processing_busy or not valid_monitors_data or not any(monitor_states): return
    processing_busy = True
    last_set_brightness = 0
    try:
        reference_brightness = next((data['cached_brightness']  
                                     for i, data in enumerate(valid_monitors_data)  
                                     if monitor_states[i]), 0)
        new_val = max(0, min(100, reference_brightness + delta))
        last_set_brightness = new_val
        for i, data in enumerate(valid_monitors_data):
            if monitor_states[i]:
                try:
                    with data['monitor_obj'] as monitor:
                        monitor.set_luminance(new_val)
                    data['cached_brightness'] = new_val
                except Exception:
                    continue
    finally:
        processing_busy = False
    show_brightness(last_set_brightness)

def increase_brightness(): threading.Thread(target=lambda: change_brightness_thread(STEP)).start()
def decrease_brightness(): threading.Thread(target=lambda: change_brightness_thread(-STEP)).start()

# --- TRAY ICON ---
def create_tray_icon_image(is_dark_mode):
    icon_color = '#ffffff' if is_dark_mode else '#000000'
    image = Image.new('RGBA', (64, 64), color=(0, 0, 0, 0))
    draw = ImageDraw.Draw(image)
    center, inner_radius, outer_radius = 32, 16, 30
    draw.ellipse([center - inner_radius, center - inner_radius, center + inner_radius, center + inner_radius], outline=icon_color, width=4)
    for angle in range(0, 360, 45):
        rad = math.radians(angle)
        x1 = center + math.cos(rad) * (inner_radius + 4)
        y1 = center + math.sin(rad) * (inner_radius + 4)
        x2 = center + math.cos(rad) * outer_radius
        y2 = center + math.sin(rad) * outer_radius
        draw.line([(x1, y1), (x2, y2)], fill=icon_color, width=4)
    return image

def on_monitor_toggle(index):
    def inner(icon, item):
        monitor_states[index] = not monitor_states[index]
        tray_icon.update_menu()
    return inner

def on_quit(icon, item):
    keyboard.unhook_all()
    if tray_icon:  
        tray_icon.stop()
    def perform_shutdown():
        global hide_timer
        if hide_timer:  
            root.after_cancel(hide_timer)
        if root.winfo_exists():
            root.quit()
            # --- SAVE SETTINGS ON EXIT ---
            try:
                settings["hotkeys"] = hotkeys
                settings["monitor_states"] = monitor_states
                with open(settings_file, "w", encoding="utf-8") as f:
                    json.dump(settings, f, indent=4)
            except Exception as e:
                print(f"Error saving settings.json: {e}")
    root.after(0, perform_shutdown)

def setup_tray():
    global tray_icon
    monitor_items = []
    
    if not valid_monitors_data:
        monitor_items.append(pystray.MenuItem("No DDC monitors found", on_quit, enabled=False))
    else:
        for i, data in enumerate(valid_monitors_data):
            name = data.get('friendly_name', f"Monitor {i + 1} (Unknown Name)")
            
            monitor_items.append(
                pystray.MenuItem(
                    name,
                    on_monitor_toggle(i),
                    checked=lambda item, idx=i: monitor_states[idx]  
                )
            )
            
    menu_items = monitor_items + [
        pystray.Menu.SEPARATOR,
        pystray.MenuItem("Open settings.json", lambda icon, item: open_settings_once()),
        pystray.MenuItem("Exit", on_quit)
    ]
    image = create_tray_icon_image(is_system_dark_mode())
    tray_icon = pystray.Icon("brightness_control", image, menu=pystray.Menu(*menu_items), title="Display Brightness")
    
    try:
        tray_icon.run()
    except Exception as e:
        print(f"Error running tray icon: {e}")


def open_settings_once():
    global hotkeys, monitor_states, settings

    if not os.path.exists(settings_file):
        print("Settings do not exist.")
        return

    proc = subprocess.Popen(["notepad.exe", settings_file])
    proc.wait()  # wait until Notepad closes

    try:
        with open(settings_file, "r", encoding="utf-8") as f:
            new_settings = json.load(f)

        hotkeys = new_settings.get("hotkeys", hotkeys)

        keyboard.unhook_all()
        bind_hotkeys()

        new_monitor_states = new_settings.get("monitor_states", [])
        for i in range(min(len(valid_monitors_data), len(new_monitor_states))):
            monitor_states[i] = new_monitor_states[i]

        settings.update(new_settings)

        print("--------------------------------------------------")
        print("Settings were loaded after closing the editor.")
        print("\nCurrent keyboard shortcuts:")
        for action, data in hotkeys.items():
            key = data.get("key")
            modifier = data.get("modifier")
            if modifier:
                print(f"{action}: {modifier} + {key}")
            else:
                print(f"{action}: {key}")

        print("\nMonitor states:")
        for i, state in enumerate(monitor_states):
            name = valid_monitors_data[i]['friendly_name'] if i < len(valid_monitors_data) else f"Monitor {i+1}"
            status = "Enabled" if state else "Disabled"
            print(f"{name}: {status}")

    except Exception as e:
        print(f"Error loading settings.json: {e}")


# --- GUI INITIALIZATION ---
root = tk.Tk()
root.overrideredirect(True)
root.attributes("-topmost", True, "-alpha", 0.98)

window_width, window_height = 280, 50
screen_width, screen_height = root.winfo_screenwidth(), root.winfo_screenheight()
x_pos = (screen_width - window_width) // 2
y_pos = int(screen_height * 0.94) - window_height
root.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")
root.withdraw()

main_frame = tk.Frame(root, padx=10, pady=7)
main_frame.pack(fill=tk.BOTH, expand=True)
content_frame = tk.Frame(main_frame)
content_frame.pack(fill=tk.BOTH, expand=True)

sun_icon = tk.Label(content_frame, text="☀️", font=("Arial", 18))
sun_icon.pack(side=tk.LEFT, padx=(0, 8))

progress_canvas = tk.Canvas(content_frame, width=200, height=6, highlightthickness=0, borderwidth=0)
progress_canvas.pack(side=tk.LEFT, padx=(0, 8))

percent_label = tk.Label(content_frame, text="0%", font=("Segoe UI", 11, "bold"), width=4, anchor='e')
percent_label.pack(side=tk.LEFT)

update_theme_colors()

# --- MAIN EXECUTION ---
#keyboard.on_press_key("up", lambda _: increase_brightness() if keyboard.is_pressed("ctrl") else None)
#keyboard.on_press_key("down", lambda _: decrease_brightness() if keyboard.is_pressed("ctrl") else None)
# --- BIND HOTKEYS FROM JSON ---
def bind_hotkeys():
    keyboard.unhook_all()  # unhook all previous binds

    for action, data in hotkeys.items():
        key = data.get("key")
        modifier = data.get("modifier")  # can be None

        if not key:
            continue

        if action == "increase_brightness":
            keyboard.on_press_key(
                key,
                lambda e, mod=modifier: increase_brightness() 
                if (mod and keyboard.is_pressed(mod)) 
                else increase_brightness() if not mod else None
            )
        elif action == "decrease_brightness":
            keyboard.on_press_key(
                key,
                lambda e, mod=modifier: decrease_brightness() 
                if (mod and keyboard.is_pressed(mod)) 
                else decrease_brightness() if not mod else None
            )


bind_hotkeys()

print("--------------------------------------------------")
print(f"Number of supported monitors: {len(valid_monitors_data)}")
if valid_monitors_data:
    names = [d['friendly_name'] for d in valid_monitors_data]
    print(f"Paired monitor names: {', '.join(names)}")

print("--------------------------------------------------")
print("Current keyboard shortcuts:")
for action, data in hotkeys.items():
    key = data.get("key")
    modifier = data.get("modifier")
    if modifier:
        print(f"{action}: {modifier} + {key}")
    else:
        print(f"{action}: {key}")


tray_thread = threading.Thread(target=setup_tray, daemon=True)
tray_thread.start()

threading.Thread(target=poll_theme_changes, daemon=True).start()

try:
    root.mainloop()
except Exception as e:
    print(f"Error in mainloop: {e}")
finally:
    on_quit(None, None)