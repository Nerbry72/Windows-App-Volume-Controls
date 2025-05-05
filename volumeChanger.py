import keyboard
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
import comtypes
import ctypes
from comtypes import CLSCTX_ALL, CoInitialize
import threading
import time
import configparser
import os
import sys
import winreg
from PIL import Image
import pystray
import tkinter as tk
from tkinter import simpledialog
import logging

# Setup logging
LOG_FILE = os.path.join(os.path.dirname(__file__), 'volume_control_debug.log')
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

logging.info("Application started.")

# Configuration setup
CONFIG_FILE = os.path.join(os.path.expanduser('~'), 'volume_control_config.ini')
config = configparser.ConfigParser()
config.read_dict({
    'Settings': {
        'program_name': 'TIDALPlayer.exe',
        'volume_steps': '5',
        'start_at_boot': 'False',
        'vol_up_shortcut': 'alt gr+up',
        'vol_down_shortcut': 'alt gr+down',
        'last_volume': '0.25'
    }
})
if os.path.exists(CONFIG_FILE):
    config.read(CONFIG_FILE)

def save_config():
    with open(CONFIG_FILE, 'w') as f:
        config.write(f)
    logging.info("Configuration saved.")

last_volume = float(config.get('Settings', 'last_volume', fallback=0.25))

# COM initialization
comtypes.CoInitialize()

def get_program_audio_session(program_name):
    logging.debug(f"Searching for audio session of program: {program_name}")
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        if session.Process and session.Process.name().lower() == program_name.lower():
            logging.debug(f"Audio session found for program: {program_name}")
            return session
    logging.warning(f"Audio session not found for program: {program_name}")
    return None

def adjust_volume(is_increase):
    global last_volume
    CoInitialize()
    program_name = config.get('Settings', 'program_name')
    volume_steps = config.getint('Settings', 'volume_steps', fallback=5)
    change = volume_steps * 0.01 * (1 if is_increase else -1)
    session = get_program_audio_session(program_name)
    if session:
        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        current_volume = volume.GetMasterVolume()
        new_volume = max(0.0, min(1.0, current_volume + change))
        last_volume = new_volume
        config.set('Settings', 'last_volume', str(last_volume))
        save_config()
        volume.SetMasterVolume(last_volume, None)
        logging.info(f"Volume adjusted to {last_volume:.2f} for program: {program_name}")
    else:
        logging.warning(f"Program '{program_name}' not found.")

def enforce_volume(stop_event):
    CoInitialize()
    while not stop_event.is_set():
        program_name = config.get('Settings', 'program_name')
        session = get_program_audio_session(program_name)
        if session:
            current_last_volume = float(config.get('Settings', 'last_volume', fallback=0.25))
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            current_volume = volume.GetMasterVolume()
            if abs(current_volume - current_last_volume) > 0.01:
                volume.SetMasterVolume(current_last_volume, None)
        time.sleep(0.5)

# Hotkey management
vol_up_handler = None
vol_down_handler = None

KEY_NORMALIZATION_MAP = {
    'Control_L': 'Control',
    'Control_R': 'Control',
    'Shift_L': 'Shift',
    'Shift_R': 'Shift',
    'Alt_L': 'Alt',
    'Alt_R': 'Alt',
    'AltGr': 'AltGr',
    'Prior': 'page up',
    'Next': 'page down',
    'exclam': '1',
    'quotedbl': '2',
    'section': '3',
    'dollar': '4',
    'percent': '5',
    'ampersand': '6',
    'slash': '7',
    'parenleft': '8',
    'parenright': '9',
    'equal': '0'
}

def normalize_keys(shortcut):
    parts = shortcut.split('+')
    normalized_parts = [KEY_NORMALIZATION_MAP.get(part, part) for part in parts]
    return '+'.join(normalized_parts)

def validate_and_get_shortcut(config_key, default):
    shortcut = config.get('Settings', config_key, fallback=default)
    shortcut = normalize_keys(shortcut)

    valid_prefixes = ['Control', 'Shift', 'Alt', 'AltGr']
    parts = shortcut.split('+')

    if not any(part in valid_prefixes for part in parts):
        logging.warning(f"Invalid shortcut for {config_key}: {shortcut}. Resetting to default '{default}'.")
        shortcut = default
        config.set('Settings', config_key, default)
        save_config()
    return shortcut

def register_hotkeys():
    global vol_up_handler, vol_down_handler
    if vol_up_handler:
        keyboard.remove_hotkey(vol_up_handler)
    if vol_down_handler:
        keyboard.remove_hotkey(vol_down_handler)

    vol_up = validate_and_get_shortcut('vol_up_shortcut', 'alt gr+up')
    vol_down = validate_and_get_shortcut('vol_down_shortcut', 'alt gr+down')

    vol_up_handler = keyboard.add_hotkey(vol_up, lambda: adjust_volume(True))
    vol_down_handler = keyboard.add_hotkey(vol_down, lambda: adjust_volume(False))
    logging.info(f"Hotkeys registered: Volume Up - {vol_up}, Volume Down - {vol_down}")

# System tray app
enforce_thread = None
stop_event = threading.Event()

def create_image():
    icon_path = os.path.join(os.path.dirname(__file__), 'volume.png')
    return Image.open(icon_path)

def on_quit(icon):
    stop_event.set()
    if enforce_thread and enforce_thread.is_alive():
        enforce_thread.join()
    icon.stop()

def on_start_at_boot(icon, item):
    new_value = not config.getboolean('Settings', 'start_at_boot', fallback=False)
    config.set('Settings', 'start_at_boot', str(new_value))
    save_config()
    set_start_at_boot(new_value)

def set_start_at_boot(enabled):
    key = winreg.HKEY_CURRENT_USER
    path = r"Software\Microsoft\Windows\CurrentVersion\Run"
    app_name = "VolumeControlApp"
    try:
        with winreg.OpenKey(key, path, 0, winreg.KEY_WRITE) as regkey:
            if enabled:
                exe_path = sys.executable
                winreg.SetValueEx(regkey, app_name, 0, winreg.REG_SZ, exe_path)
            else:
                try:
                    winreg.DeleteValue(regkey, app_name)
                except FileNotFoundError:
                    pass
    except Exception as e:
        print(f"Error setting start at boot: {e}")

def set_shortcut(icon):
    def record_hotkey(prompt):
        root = tk.Tk()
        root.title("Set Shortcut")
        label = tk.Label(root, text=prompt)
        label.pack(pady=10)

        hotkey_var = tk.StringVar()
        entry = tk.Entry(root, textvariable=hotkey_var, state='readonly', width=30)
        entry.pack(pady=5)

        pressed_keys = []

        def on_key_press(event):
            key = event.keysym
            if key not in pressed_keys:
                pressed_keys.append(key)
            hotkey_var.set('+'.join(pressed_keys))

        def on_key_release(event):
            pass  

        def clear_hotkey():
            pressed_keys.clear()
            hotkey_var.set("")

        root.bind("<KeyPress>", on_key_press)
        root.bind("<KeyRelease>", on_key_release)

        clear_button = tk.Button(root, text="Clear Hotkey", command=clear_hotkey)
        clear_button.pack(pady=5)

        def on_ok():
            root.destroy()

        ok_button = tk.Button(root, text="OK", command=on_ok)
        ok_button.pack(pady=10)

        root.mainloop()

        shifted_keys_map = {
            'exclam': '1',
            'quotedbl': '2',
            'section': '3',
            'dollar': '4',
            'percent': '5',
            'ampersand': '6',
            'slash': '7',
            'parenleft': '8',
            'parenright': '9',
            'equal': '0'
        }
        parts = hotkey_var.get().split('+')
        fixed_parts = [shifted_keys_map.get(part, part) for part in parts]
        return '+'.join(fixed_parts)

    def thread_func():
        new_up = record_hotkey("Press hotkey combination or click 'Clear Hotkey' to remove the current hotkey for Volume Up")
        new_down = record_hotkey("Press hotkey combination or click 'Clear Hotkey' to remove the current hotkey for Volume Down")
        if new_up and new_down:
            config.set('Settings', 'vol_up_shortcut', new_up)
            config.set('Settings', 'vol_down_shortcut', new_down)
            save_config()
            register_hotkeys()
            icon.notify("Shortcuts updated!", "Volume Control")

    threading.Thread(target=thread_func, daemon=True).start()

def set_program(icon):
    global enforce_thread
    root = tk.Tk()
    root.withdraw()
    program_name = simpledialog.askstring("Set Program", "Enter the program's EXE name:")
    if program_name:
        config.set('Settings', 'program_name', program_name)
        save_config()
        logging.info(f"Program name set to: {program_name}")
        stop_event.set()
        if enforce_thread and enforce_thread.is_alive():
            enforce_thread.join()
        stop_event.clear()
        enforce_thread = threading.Thread(target=enforce_volume, args=(stop_event,), daemon=True)
        enforce_thread.start()

def set_volume_steps(icon):
    root = tk.Tk()
    root.withdraw()
    steps = simpledialog.askinteger("Set Volume Steps", "Enter steps (1-10):", minvalue=1, maxvalue=10)
    if steps:
        config.set('Settings', 'volume_steps', str(steps))
        save_config()
        logging.info(f"Volume steps set to: {steps}")

menu = pystray.Menu(
    pystray.MenuItem(
        'Start at boot',
        on_start_at_boot,
        checked=lambda item: config.getboolean('Settings', 'start_at_boot', fallback=False)
    ),
    pystray.MenuItem('Set Shortcut', set_shortcut),
    pystray.MenuItem('Set Program', set_program),
    pystray.MenuItem('Set Volume Steps', set_volume_steps),
    pystray.MenuItem('Quit', on_quit)
)

icon = pystray.Icon("Volume Control", create_image(), menu=menu)

register_hotkeys()
enforce_thread = threading.Thread(target=enforce_volume, args=(stop_event,), daemon=True)
enforce_thread.start()
set_start_at_boot(config.getboolean('Settings', 'start_at_boot', fallback=False))

icon.run()