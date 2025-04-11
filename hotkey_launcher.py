import sys
import os
from PyQt5.QtWidgets import QApplication
from PyQt5.QtCore import QTimer
from ui import FloatingChat
import keyboard

assistant_window = None
was_pressed = False  # track toggle state


def launch_ui():
    global assistant_window
    if assistant_window is None:
        print("[HOTKEY] Launching assistant window...")
        assistant_window = FloatingChat()
        assistant_window.destroyed.connect(clear_assistant_ref)
        assistant_window.show()
        QTimer.singleShot(1000, assistant_window.initialize_context)
    else:
        print("[HOTKEY] Assistant already open.")


def clear_assistant_ref():
    global assistant_window
    print("[INFO] Assistant window closed.")
    assistant_window = None


def check_hotkey():
    global was_pressed
    if keyboard.is_pressed("ctrl+alt+a"):
        if not was_pressed:
            was_pressed = True
            launch_ui()
    else:
        was_pressed = False

    QTimer.singleShot(100, check_hotkey)  # repeat every 100ms


if __name__ == "__main__":
    app = QApplication(sys.argv)
    print("🔑 Hold Ctrl+Alt+A to launch your assistant.")
    check_hotkey()
    sys.exit(app.exec_())
