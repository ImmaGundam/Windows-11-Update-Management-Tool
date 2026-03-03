import pystray
from PIL import Image, ImageDraw
import threading
import subprocess
import time
import webview
import ctypes
import winreg
import sys
import os

# Prevent subprocess console windows
CREATE_NO_WINDOW = 0x08000000

# -------------------------
# CONFIG
# -------------------------

NORMAL_SERVICES = ["wuauserv", "UsoSvc"]
DEFENDER_SERVICES = ["WinDefend", "WdNisSvc", "Sense"]

ignore_defender = False
icon = None


# -------------------------
# ADMIN CHECK / ELEVATION
# -------------------------

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False


def elevate():
    """Relaunch with admin privileges"""
    if sys.argv[0].endswith(".pyw") or sys.argv[0].endswith(".py"):
        ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            sys.executable,
            " ".join(sys.argv),
            None,
            1
        )
    else:
        ctypes.windll.shell32.ShellExecuteW(
            None,
            "runas",
            sys.argv[0],
            None,
            None,
            1
        )


def require_admin():
    global icon

    if not is_admin():

        result = ctypes.windll.user32.MessageBoxW(
            0,
            "This action requires Administrator privileges.\n\nClick YES to restart as Administrator.",
            "Administrator Required",
            0x04 | 0x20
        )

        if result == 6:  # YES

            # Stop tray cleanly
            if icon is not None:
                try:
                    icon.stop()
                except:
                    pass

            elevate()

            # Hard exit to prevent duplicate instance
            os._exit(0)

        return False

    return True


# -------------------------
# SERVICE UTILITIES
# -------------------------

def get_service_state(service):
    try:
        result = subprocess.check_output(
            ["sc", "query", service],
            text=True,
            stderr=subprocess.DEVNULL,
            creationflags=CREATE_NO_WINDOW
        )

        if "RUNNING" in result:
            return "Running"
        elif "STOPPED" in result:
            return "Stopped"
        else:
            return "Unknown"

    except:
        return "Unknown"


def stop_and_disable_service(service):
    subprocess.run(
        ["sc", "stop", service],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=CREATE_NO_WINDOW
    )
    subprocess.run(
        ["sc", "config", service, "start=", "disabled"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=CREATE_NO_WINDOW
    )


def set_auto_and_start(service):
    subprocess.run(
        ["sc", "config", service, "start=", "auto"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=CREATE_NO_WINDOW
    )
    subprocess.run(
        ["sc", "start", service],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=CREATE_NO_WINDOW
    )


# -------------------------
# DEFENDER FIX
# -------------------------

def remove_defender_policy_blocks():
    policy_path = r"SOFTWARE\Policies\Microsoft\Windows Defender"
    try:
        winreg.DeleteKey(winreg.HKEY_LOCAL_MACHINE, policy_path)
    except:
        pass


def fix_defender_logic():
    if not require_admin():
        return

    remove_defender_policy_blocks()

    for svc in DEFENDER_SERVICES:
        set_auto_and_start(svc)


def run_all_logic():
    if not require_admin():
        return

    for svc in NORMAL_SERVICES:
        stop_and_disable_service(svc)


# -------------------------
# SERVICE STATUS
# -------------------------

def check_services():
    normal_ok = True

    for svc in NORMAL_SERVICES:
        if get_service_state(svc) == "Running":
            normal_ok = False

    if ignore_defender:
        defender_ok = None
    else:
        defender_ok = True
        for svc in DEFENDER_SERVICES:
            if get_service_state(svc) != "Running":
                defender_ok = False

    return normal_ok, defender_ok


# -------------------------
# API FOR HTML
# -------------------------

class API:

    def get_status(self):
        normal, defender = check_services()
        return {"services": normal, "defender": defender}

    def run_all(self):
        run_all_logic()
        return self.get_status()

    def fix_defender(self):
        fix_defender_logic()
        return self.get_status()

    def toggle_ignore(self, value):
        global ignore_defender
        ignore_defender = value
        return self.get_status()


# -------------------------
# TRAY ICON
# -------------------------

def create_icon(color):
    img = Image.new("RGB", (16, 16), (0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((2, 2, 14, 14), fill=color)
    return img


def update_tray():
    global icon

    while True:
        normal, defender = check_services()

        if normal and (defender or defender is None):
            icon.icon = create_icon("green")
        else:
            icon.icon = create_icon("red")

        time.sleep(10)


def open_panel(icon, item):
    webview.windows[0].show()


def exit_app(icon, item):
    icon.stop()
    webview.windows[0].destroy()


def on_closing():
    webview.windows[0].hide()
    return False  # Prevent actual close


def run_tray():
    global icon

    icon = pystray.Icon(
        "Watchdog",
        create_icon("gray"),
        "Windows Update Watchdog",
        menu=pystray.Menu(
            pystray.MenuItem("Open", open_panel),
            pystray.MenuItem("Exit", exit_app)
        )
    )

    threading.Thread(target=update_tray, daemon=True).start()
    icon.run()


# -------------------------
# MAIN
# -------------------------

threading.Thread(target=run_tray, daemon=True).start()

window = webview.create_window(
    "Windows Update Watchdog",
    "ui.html",
    js_api=API(),
    width=200,
    height=130,
    hidden=True,
    resizable=False
)

window.events.closing += on_closing

webview.start()