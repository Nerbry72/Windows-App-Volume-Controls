import keyboard
from pycaw.pycaw import AudioUtilities, ISimpleAudioVolume
from comtypes import CLSCTX_ALL
import ctypes
import comtypes
from comtypes import CoInitialize
import threading
import time

comtypes.CoInitialize()


last_volume = 0.25

def get_program_audio_session(program_name):
    sessions = AudioUtilities.GetAllSessions()
    for session in sessions:
        if session.Process and session.Process.name().lower() == program_name.lower():
            return session
    return None


def adjust_volume(program_name, change):
    global last_volume
    CoInitialize()  
    session = get_program_audio_session(program_name)
    if session:

        volume = session._ctl.QueryInterface(ISimpleAudioVolume)
        current_volume = volume.GetMasterVolume()
        last_volume = max(0.0, min(1.0, current_volume + change))
        volume.SetMasterVolume(last_volume, None)

        print(f"Volume for {program_name} set to {last_volume * 100:.0f}%")
    else:
        print(f"Program '{program_name}' not found.")


def enforce_volume(program_name):
    CoInitialize() 
    global last_volume
    while True:
        session = get_program_audio_session(program_name)
        if session:
            volume = session._ctl.QueryInterface(ISimpleAudioVolume)
            current_volume = volume.GetMasterVolume()
            if abs(current_volume - last_volume) > 0.01:  
                volume.SetMasterVolume(last_volume, None)
        time.sleep(1)  

PROGRAM_NAME = "TIDALPlayer.exe" 

keyboard.add_hotkey('alt gr+up', lambda: adjust_volume(PROGRAM_NAME, 0.05))
keyboard.add_hotkey('alt gr+down', lambda: adjust_volume(PROGRAM_NAME, -0.05))

threading.Thread(target=enforce_volume, args=(PROGRAM_NAME,), daemon=True).start()

print("Hotkeys registered. Use Alt Gr + Up to increase volume and Alt Gr + Down to decrease volume.")
print("Press ESC to exit.")

keyboard.wait('esc')