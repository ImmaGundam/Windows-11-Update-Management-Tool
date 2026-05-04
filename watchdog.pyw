import ctypes
import ctypes.wintypes
import json
import os
import re
import subprocess
import sys
import threading
import time
import winreg

import pystray
import webview
from PIL import Image, ImageDraw

# Prevent subprocess console windows on Windows
CREATE_NO_WINDOW = 0x08000000

APP_NAME = "Windows Update Watchdog"
APP_VERSION = "1.1.0"
APP_ID = "WebGeeksIT.WindowsUpdateWatchdog"
APP_ICON_RELATIVE_PATH = os.path.join("assets", "app_icon.ico")

# -------------------------
# CONFIG
# -------------------------

# Strictly monitored update services. These should be stopped + disabled after Run All.
UPDATE_SERVICES = ["wuauserv", "bits", "dosvc", "UsoSvc"]

# Attempted during Run All, but not used as a hard red/green condition because
# Windows can protect or recreate this service depending on build/policy state.
UPDATE_OPTIONAL_SERVICES = ["WaaSMedicSvc"]

# Defender service model:
# - WinDefend should be running.
# - WdNisSvc should exist and not be disabled, but may be stopped by Windows until needed.
# - Sense is Microsoft Defender for Endpoint and is optional on home/small-business PCs.
DEFENDER_CORE_SERVICE = "WinDefend"
DEFENDER_SUPPORT_SERVICES = ["WdNisSvc"]
DEFENDER_OPTIONAL_SERVICES = ["Sense"]

ignore_defender = False
update_guard_enabled = False
update_guard_interval_seconds = 8
icon = None
window = None
status_window = None
api_instance = None
status_lock = threading.Lock()
activity_lock = threading.Lock()
activity_log = []
current_action = "Idle"
shutdown_requested = threading.Event()
MAX_LOG_LINES = 250

# PowerShell Defender status is heavier than service checks, so cache it briefly.
defender_cache = {
    "checked_at": 0.0,
    "status": None,
}


# -------------------------
# PATH / PROCESS HELPERS
# -------------------------

def resource_path(relative_path):
    """Resolve resource paths correctly in source and PyInstaller builds."""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(os.path.dirname(__file__))
    return os.path.join(base_path, relative_path)


def app_icon_path():
    """Return the app icon path used by pywebview, Windows, and PyInstaller."""
    return resource_path(APP_ICON_RELATIVE_PATH)


def set_windows_app_id():
    """Set a Windows AppUserModelID so the taskbar groups under this app name/icon."""
    if os.name != "nt":
        return

    try:
        ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(APP_ID)
    except Exception:
        pass


def apply_native_window_icons(*_args):
    """Apply the bundled ICO to pywebview top-level Windows title bars.

    PyInstaller sets the executable icon, but pywebview windows can still show
    the default Python/window icon when running from source or before Windows
    refreshes the title bar. This explicitly assigns the ICO to the main and
    diagnostics windows.
    """
    if os.name != "nt":
        return

    icon_file = app_icon_path()
    if not os.path.exists(icon_file):
        return

    try:
        user32 = ctypes.windll.user32
        kernel32 = ctypes.windll.kernel32

        IMAGE_ICON = 1
        LR_LOADFROMFILE = 0x00000010
        LR_DEFAULTSIZE = 0x00000040
        WM_SETICON = 0x0080
        ICON_SMALL = 0
        ICON_BIG = 1
        GCLP_HICON = -14
        GCLP_HICONSM = -34

        hicon_big = user32.LoadImageW(None, icon_file, IMAGE_ICON, 0, 0, LR_LOADFROMFILE | LR_DEFAULTSIZE)
        hicon_small = user32.LoadImageW(None, icon_file, IMAGE_ICON, 16, 16, LR_LOADFROMFILE)

        if not hicon_big:
            return

        current_pid = os.getpid()
        target_titles = {APP_NAME, "Watchdog Status"}

        EnumWindowsProc = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)

        def enum_callback(hwnd, lparam):
            pid = ctypes.wintypes.DWORD()
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))

            if pid.value != current_pid:
                return True

            length = user32.GetWindowTextLengthW(hwnd)
            buffer = ctypes.create_unicode_buffer(length + 1)
            user32.GetWindowTextW(hwnd, buffer, length + 1)
            title = buffer.value

            if title in target_titles or title.startswith(APP_NAME):
                user32.SendMessageW(hwnd, WM_SETICON, ICON_BIG, hicon_big)
                user32.SendMessageW(hwnd, WM_SETICON, ICON_SMALL, hicon_small or hicon_big)

                if hasattr(user32, "SetClassLongPtrW"):
                    user32.SetClassLongPtrW(hwnd, GCLP_HICON, hicon_big)
                    user32.SetClassLongPtrW(hwnd, GCLP_HICONSM, hicon_small or hicon_big)
                else:
                    user32.SetClassLongW(hwnd, GCLP_HICON, hicon_big)
                    user32.SetClassLongW(hwnd, GCLP_HICONSM, hicon_small or hicon_big)

            return True

        user32.EnumWindows(EnumWindowsProc(enum_callback), 0)

    except Exception:
        pass


def subprocess_kwargs(timeout=15):
    kwargs = {
        "capture_output": True,
        "text": True,
        "timeout": timeout,
    }
    if os.name == "nt":
        kwargs["creationflags"] = CREATE_NO_WINDOW
    return kwargs


def run_cmd(args, timeout=15):
    """Run a command and return structured success/error details."""
    try:
        result = subprocess.run(args, **subprocess_kwargs(timeout=timeout))
        return {
            "ok": result.returncode == 0,
            "code": result.returncode,
            "stdout": (result.stdout or "").strip(),
            "stderr": (result.stderr or "").strip(),
            "cmd": " ".join(args),
        }
    except Exception as exc:
        return {
            "ok": False,
            "code": None,
            "stdout": "",
            "stderr": str(exc),
            "cmd": " ".join(args),
        }


def run_powershell(script, timeout=25):
    return run_cmd(
        [
            "powershell",
            "-NoProfile",
            "-ExecutionPolicy",
            "Bypass",
            "-Command",
            script,
        ],
        timeout=timeout,
    )


# -------------------------
# ACTIVITY LOG
# -------------------------

def add_log(message):
    """Append a short terminal-style activity line for the diagnostics window."""
    timestamp = time.strftime("%H:%M:%S")
    line = f"[{timestamp}] {message}"

    with activity_lock:
        activity_log.append(line)
        if len(activity_log) > MAX_LOG_LINES:
            del activity_log[:-MAX_LOG_LINES]


def set_current_action(action):
    global current_action
    with activity_lock:
        current_action = action


def get_activity_snapshot():
    with activity_lock:
        return {
            "current_action": current_action,
            "lines": list(activity_log[-MAX_LOG_LINES:]),
        }


def log_service_change(prefix, item):
    service = item.get("service", "unknown")

    if item.get("skipped"):
        add_log(f"{prefix}: {service} skipped ({item.get('reason', 'not applicable')})")
        return

    before = item.get("before") or {}
    after = item.get("after") or {}
    before_text = f"{before.get('state', '?')}/{before.get('start_type', '?')}"
    after_text = f"{after.get('state', '?')}/{after.get('start_type', '?')}"
    result = "OK" if item.get("ok") else "CHECK"
    add_log(f"{prefix}: {service} {before_text} -> {after_text} [{result}]")


def log_update_results(prefix, results):
    for item in results.get("services", []):
        log_service_change(prefix, item)
    for item in results.get("optional_services", []):
        log_service_change(prefix + " optional", item)

    policy = results.get("policy")
    if isinstance(policy, dict):
        add_log(f"{prefix}: update policy {'OK' if policy.get('ok') else 'CHECK'}")

    tasks = results.get("tasks")
    if isinstance(tasks, dict):
        add_log(f"{prefix}: UpdateOrchestrator tasks {'OK' if tasks.get('ok') else 'CHECK'}")

    metered = results.get("metered_ethernet")
    if isinstance(metered, dict):
        add_log(f"{prefix}: Ethernet metered setting {'OK' if metered.get('ok') else 'CHECK'}")


def log_defender_results(prefix, results):
    for item in results.get("services", []):
        log_service_change(prefix, item)

    removed = results.get("removed_policies", [])
    removed_count = sum(1 for item in removed if item.get("removed"))
    add_log(f"{prefix}: Defender policy cleanup removed {removed_count} value(s)")

    sig = results.get("signature_updates")
    if isinstance(sig, dict):
        add_log(f"{prefix}: Defender signature fallback {'OK' if sig.get('ok') else 'CHECK'}")

    add_log(f"{prefix}: Defender status {'OK' if results.get('ok') else 'CHECK'}")


# -------------------------
# ADMIN / ELEVATION
# -------------------------

def is_admin():
    if os.name != "nt":
        return False
    try:
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return False


def elevate():
    """Relaunch with admin privileges."""
    if getattr(sys, "frozen", False):
        file_to_run = sys.executable
        params = ""
    else:
        file_to_run = sys.executable
        params = subprocess.list2cmdline(sys.argv)

    ctypes.windll.shell32.ShellExecuteW(
        None,
        "runas",
        file_to_run,
        params,
        None,
        1,
    )


def require_admin():
    global icon

    if is_admin():
        return True

    add_log("Administrator elevation requested")

    result = ctypes.windll.user32.MessageBoxW(
        0,
        "This action requires Administrator privileges.\n\nClick YES to restart as Administrator.",
        "Administrator Required",
        0x04 | 0x20,
    )

    if result == 6:  # YES
        if icon is not None:
            try:
                icon.stop()
            except Exception:
                pass
        elevate()
        os._exit(0)

    add_log("Administrator elevation declined")
    return False


# -------------------------
# SERVICE UTILITIES
# -------------------------

def get_service_info(service):
    """Return service existence, state, and startup type using sc.exe."""
    query = run_cmd(["sc", "query", service], timeout=8)
    qc = run_cmd(["sc", "qc", service], timeout=8)

    if not query["ok"] and not qc["ok"]:
        return {
            "name": service,
            "exists": False,
            "state": "Missing",
            "start_type": "Missing",
            "query": query,
            "qc": qc,
        }

    state = "Unknown"
    state_match = re.search(r"STATE\s*:\s*\d+\s+(\w+)", query["stdout"])
    if state_match:
        raw_state = state_match.group(1).upper()
        if raw_state == "RUNNING":
            state = "Running"
        elif raw_state == "STOPPED":
            state = "Stopped"
        else:
            state = raw_state.title()

    start_type = "Unknown"
    start_match = re.search(r"START_TYPE\s*:\s*\d+\s+(\w+)", qc["stdout"])
    if start_match:
        raw_start = start_match.group(1).upper()
        if raw_start == "DISABLED":
            start_type = "Disabled"
        elif raw_start in ("DEMAND_START", "DEMAND"):
            start_type = "Manual"
        elif raw_start in ("AUTO_START", "AUTO"):
            start_type = "Automatic"
        elif raw_start == "BOOT_START":
            start_type = "Boot"
        elif raw_start == "SYSTEM_START":
            start_type = "System"
        else:
            start_type = raw_start.title()

    return {
        "name": service,
        "exists": True,
        "state": state,
        "start_type": start_type,
        "query": query,
        "qc": qc,
    }


def stop_service(service):
    return run_cmd(["sc", "stop", service], timeout=15)


def start_service(service):
    return run_cmd(["sc", "start", service], timeout=15)


def set_service_start_type(service, start_type):
    # start_type: disabled, demand, auto
    return run_cmd(["sc", "config", service, "start=", start_type], timeout=15)


def set_service_registry_start(service, value):
    # HKLM\SYSTEM\CurrentControlSet\Services\<service>\Start
    # 2 = Automatic, 3 = Manual, 4 = Disabled
    return set_dword_hklm(
        rf"SYSTEM\CurrentControlSet\Services\{service}",
        "Start",
        value,
    )


def powershell_stop_and_disable_service(service):
    # PowerShell fallback mirrors the original script and helps when sc.exe gives
    # incomplete results or Windows flips a service state during the operation.
    script = rf'''
    try {{
        Stop-Service -Name "{service}" -Force -ErrorAction SilentlyContinue
        Set-Service -Name "{service}" -StartupType Disabled -ErrorAction SilentlyContinue
        "ok"
    }} catch {{
        "error: $($_.Exception.Message)"
    }}
    '''
    return run_powershell(script, timeout=20)


def update_service_required_ok(svc):
    if not svc or not svc.get("exists"):
        return True
    return svc.get("state") != "Running" and svc.get("start_type") == "Disabled"


def update_service_optional_ok(svc):
    if not svc or not svc.get("exists"):
        return True
    # WaaSMedicSvc is protected on many Windows builds. Treat it as a warning
    # only when active; stopped/manual is informational.
    return svc.get("state") != "Running"


def stop_and_disable_service(service):
    before = get_service_info(service)
    results = []

    if not before["exists"]:
        return {"service": service, "ok": True, "skipped": True, "reason": "missing", "before": before}

    if before["state"] == "Running":
        results.append(stop_service(service))
        results.append(powershell_stop_and_disable_service(service))

    # Use both service manager and registry startup settings. Windows Update can
    # re-arm services quickly, so the guard loop may need to repeat this later.
    results.append(set_service_start_type(service, "disabled"))
    results.append(set_service_registry_start(service, 4))

    after_first = get_service_info(service)
    if after_first["state"] == "Running":
        results.append(stop_service(service))
        results.append(powershell_stop_and_disable_service(service))

    after = get_service_info(service)

    return {
        "service": service,
        "ok": update_service_required_ok(after),
        "before": before,
        "after": after,
        "results": results,
    }


def enable_service(service, start_type="demand", should_start=True):
    before = get_service_info(service)
    results = []

    if not before["exists"]:
        return {"service": service, "ok": True, "skipped": True, "reason": "missing", "before": before}

    start_value = 2 if start_type == "auto" else 3
    results.append(set_service_registry_start(service, start_value))
    results.append(set_service_start_type(service, start_type))
    if should_start:
        results.append(start_service(service))

    after = get_service_info(service)
    return {
        "service": service,
        "ok": after["exists"] and after["start_type"] != "Disabled",
        "before": before,
        "after": after,
        "results": results,
    }


# -------------------------
# REGISTRY HELPERS
# -------------------------

def open_or_create_hklm_key(path, access):
    return winreg.CreateKeyEx(
        winreg.HKEY_LOCAL_MACHINE,
        path,
        0,
        access | winreg.KEY_WOW64_64KEY,
    )


def set_dword_hklm(path, name, value):
    try:
        with open_or_create_hklm_key(path, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, name, 0, winreg.REG_DWORD, int(value))
        return {"ok": True, "path": path, "name": name, "value": value}
    except Exception as exc:
        return {"ok": False, "path": path, "name": name, "error": str(exc)}


def set_string_hklm(path, name, value):
    try:
        with open_or_create_hklm_key(path, winreg.KEY_SET_VALUE) as key:
            winreg.SetValueEx(key, name, 0, winreg.REG_SZ, str(value))
        return {"ok": True, "path": path, "name": name, "value": value}
    except Exception as exc:
        return {"ok": False, "path": path, "name": name, "error": str(exc)}


def read_value_hklm(path, name):
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            path,
            0,
            winreg.KEY_READ | winreg.KEY_WOW64_64KEY,
        ) as key:
            value, value_type = winreg.QueryValueEx(key, name)
        return {"exists": True, "value": value, "type": value_type}
    except FileNotFoundError:
        return {"exists": False, "value": None, "type": None}
    except Exception as exc:
        return {"exists": False, "value": None, "type": None, "error": str(exc)}


def delete_value_hklm(path, name):
    try:
        with winreg.OpenKey(
            winreg.HKEY_LOCAL_MACHINE,
            path,
            0,
            winreg.KEY_SET_VALUE | winreg.KEY_WOW64_64KEY,
        ) as key:
            winreg.DeleteValue(key, name)
        return {"ok": True, "path": path, "name": name, "removed": True}
    except FileNotFoundError:
        return {"ok": True, "path": path, "name": name, "removed": False, "reason": "not found"}
    except Exception as exc:
        return {"ok": False, "path": path, "name": name, "error": str(exc)}


# -------------------------
# WINDOWS UPDATE CONTROLS
# -------------------------

def set_windows_update_policy_disabled():
    return set_dword_hklm(
        r"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU",
        "NoAutoUpdate",
        1,
    )


def remove_windows_update_policy_disabled():
    return delete_value_hklm(
        r"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU",
        "NoAutoUpdate",
    )


def is_windows_update_policy_disabled():
    value = read_value_hklm(
        r"SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU",
        "NoAutoUpdate",
    )
    return value.get("exists") and int(value.get("value") or 0) == 1


def disable_update_orchestrator_tasks():
    script = r"""
    Get-ScheduledTask | Where-Object {
        $_.TaskPath -like '\Microsoft\Windows\UpdateOrchestrator*'
    } | ForEach-Object {
        Disable-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction SilentlyContinue | Out-Null
    }
    "ok"
    """
    return run_powershell(script, timeout=30)


def enable_update_orchestrator_tasks():
    script = r"""
    Get-ScheduledTask | Where-Object {
        $_.TaskPath -like '\Microsoft\Windows\UpdateOrchestrator*'
    } | ForEach-Object {
        Enable-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -ErrorAction SilentlyContinue | Out-Null
    }
    "ok"
    """
    return run_powershell(script, timeout=30)


def set_ethernet_metered(enabled):
    # Mirrors the original PowerShell behavior. This may fail on some Windows builds
    # due to DefaultMediaCost key permissions; failure is reported but not fatal.
    value = 2 if enabled else 1
    return set_dword_hklm(
        r"SOFTWARE\Microsoft\Windows NT\CurrentVersion\NetworkList\DefaultMediaCost",
        "Ethernet",
        value,
    )


def disable_update_controls():
    results = {
        "services": [],
        "optional_services": [],
        "policy": None,
        "tasks": None,
        "metered_ethernet": None,
    }

    for service in UPDATE_SERVICES:
        results["services"].append(stop_and_disable_service(service))

    for service in UPDATE_OPTIONAL_SERVICES:
        item = stop_and_disable_service(service)
        after = item.get("after") or item.get("before")
        item["ok"] = update_service_optional_ok(after)
        results["optional_services"].append(item)

    results["policy"] = set_windows_update_policy_disabled()
    results["tasks"] = disable_update_orchestrator_tasks()
    results["metered_ethernet"] = set_ethernet_metered(True)
    return results


def restore_update_controls():
    results = {
        "services": [],
        "optional_services": [],
        "policy": None,
        "tasks": None,
        "metered_ethernet": None,
        "defender_signature_policy": None,
    }

    # Original script restored services to Manual and started them.
    for service in UPDATE_SERVICES:
        results["services"].append(enable_service(service, start_type="demand", should_start=True))

    for service in UPDATE_OPTIONAL_SERVICES:
        results["optional_services"].append(enable_service(service, start_type="demand", should_start=True))

    results["policy"] = remove_windows_update_policy_disabled()
    results["tasks"] = enable_update_orchestrator_tasks()
    results["metered_ethernet"] = set_ethernet_metered(False)
    results["defender_signature_policy"] = delete_value_hklm(
        r"SOFTWARE\Policies\Microsoft\Windows Defender\Signature Updates",
        "FallbackOrder",
    )
    return results


# -------------------------
# DEFENDER PROTECTION / FIX
# -------------------------

def remove_known_defender_policy_blocks():
    # Target only known disabling values instead of deleting the whole policy key.
    targets = {
        r"SOFTWARE\Policies\Microsoft\Windows Defender": [
            "DisableAntiSpyware",
            "DisableAntiVirus",
            "DisableSpecialRunningModes",
            "ServiceKeepAlive",
        ],
        r"SOFTWARE\Policies\Microsoft\Windows Defender\Real-Time Protection": [
            "DisableRealtimeMonitoring",
            "DisableBehaviorMonitoring",
            "DisableOnAccessProtection",
            "DisableScanOnRealtimeEnable",
            "DisableIOAVProtection",
        ],
        r"SOFTWARE\Policies\Microsoft\Windows Defender\Spynet": [
            "DisableBlockAtFirstSeen",
        ],
    }

    results = []
    for path, names in targets.items():
        for name in names:
            results.append(delete_value_hklm(path, name))
    return results


def enable_defender_signature_updates():
    # Original script forces Defender signatures to use Microsoft Malware Protection Center.
    return set_string_hklm(
        r"SOFTWARE\Policies\Microsoft\Windows Defender\Signature Updates",
        "FallbackOrder",
        "MMPC",
    )


def get_defender_mp_status_cached(max_age_seconds=30):
    now = time.time()
    if defender_cache["status"] is not None and now - defender_cache["checked_at"] < max_age_seconds:
        return defender_cache["status"]

    script = r"""
    try {
        $s = Get-MpComputerStatus -ErrorAction Stop
        [PSCustomObject]@{
            AMServiceEnabled = $s.AMServiceEnabled
            AntivirusEnabled = $s.AntivirusEnabled
            RealTimeProtectionEnabled = $s.RealTimeProtectionEnabled
            BehaviorMonitorEnabled = $s.BehaviorMonitorEnabled
            IoavProtectionEnabled = $s.IoavProtectionEnabled
            NISEnabled = $s.NISEnabled
            IsTamperProtected = $s.IsTamperProtected
        } | ConvertTo-Json -Compress
    } catch {
        [PSCustomObject]@{
            Error = $_.Exception.Message
        } | ConvertTo-Json -Compress
    }
    """

    result = run_powershell(script, timeout=12)
    status = {"ok": False, "source": "Get-MpComputerStatus", "raw": result}

    if result["ok"] and result["stdout"]:
        try:
            data = json.loads(result["stdout"].splitlines()[-1])
            if data.get("Error"):
                status = {"ok": False, "source": "Get-MpComputerStatus", "data": data, "raw": result}
            else:
                protection_ok = (
                    data.get("AMServiceEnabled") is True
                    and data.get("AntivirusEnabled") is True
                    and data.get("RealTimeProtectionEnabled") is True
                )
                status = {
                    "ok": protection_ok,
                    "source": "Get-MpComputerStatus",
                    "data": data,
                    "raw": result,
                }
        except Exception as exc:
            status = {"ok": False, "source": "Get-MpComputerStatus", "error": str(exc), "raw": result}

    defender_cache["checked_at"] = now
    defender_cache["status"] = status
    return status


def check_defender():
    core = get_service_info(DEFENDER_CORE_SERVICE)
    support = [get_service_info(service) for service in DEFENDER_SUPPORT_SERVICES]
    optional = [get_service_info(service) for service in DEFENDER_OPTIONAL_SERVICES]

    core_ok = core["exists"] and core["state"] == "Running" and core["start_type"] != "Disabled"
    support_ok = all((not svc["exists"]) or svc["start_type"] != "Disabled" for svc in support)

    mp_status = get_defender_mp_status_cached()

    # If Get-MpComputerStatus works, trust it for protection state.
    # If it fails, fall back to core service state so the UI does not false-red on stripped systems.
    protection_ok = mp_status["ok"] if mp_status.get("raw", {}).get("ok") else core_ok

    ok = bool(core_ok and support_ok and protection_ok)

    return {
        "ok": ok,
        "core": core,
        "support": support,
        "optional": optional,
        "mp_status": mp_status,
    }


def fix_defender_logic(check_admin=True):
    if check_admin and not require_admin():
        return {"ok": False, "admin": False}

    results = {
        "admin": True,
        "removed_policies": remove_known_defender_policy_blocks(),
        "services": [],
        "signature_updates": enable_defender_signature_updates(),
    }

    # Keep Defender AV running.
    results["services"].append(enable_service(DEFENDER_CORE_SERVICE, start_type="auto", should_start=True))

    # WdNisSvc is a support service. Keep it available, but do not require it to always be running.
    for service in DEFENDER_SUPPORT_SERVICES:
        results["services"].append(enable_service(service, start_type="demand", should_start=True))

    # Sense is optional. Enable/start only if present.
    for service in DEFENDER_OPTIONAL_SERVICES:
        info = get_service_info(service)
        if info["exists"]:
            results["services"].append(enable_service(service, start_type="demand", should_start=True))

    defender_cache["checked_at"] = 0.0
    results["after"] = check_defender()
    results["ok"] = results["after"]["ok"]
    return results


# -------------------------
# STATUS
# -------------------------

def check_update_controls():
    services = [get_service_info(service) for service in UPDATE_SERVICES]
    optional_services = [get_service_info(service) for service in UPDATE_OPTIONAL_SERVICES]

    services_ok = all(update_service_required_ok(svc) for svc in services)
    optional_ok = all(update_service_optional_ok(svc) for svc in optional_services)
    policy_ok = is_windows_update_policy_disabled()

    return {
        "ok": bool(services_ok and optional_ok and policy_ok),
        "services": services,
        "optional_services": optional_services,
        "policy_disabled": policy_ok,
        "guard_enabled": update_guard_enabled,
    }


def check_status():
    with status_lock:
        updates = check_update_controls()

        # Still collect Defender details while ignored so the diagnostics window
        # can gray the information out instead of hiding it. The summary state
        # remains None so Defender does not affect overall status.
        defender = check_defender()

        return {
            "services": updates["ok"],
            "defender": None if ignore_defender else defender["ok"],
            "details": {
                "updates": updates,
                "defender": defender,
                "ignore_defender": ignore_defender,
                "admin": is_admin(),
                "version": APP_VERSION,
                "guard_interval_seconds": update_guard_interval_seconds,
            },
        }


# -------------------------
# UPDATE GUARD
# -------------------------

def update_guard_loop():
    """Re-apply update controls after Run All if Windows re-enables them.

    Windows Update Medic, BITS jobs, Update Orchestrator, or Windows servicing
    can flip wuauserv/BITS/DoSvc back to Manual or start them again after the
    first pass. This loop makes the app a real watchdog without changing the
    tiny main UI: Start turns the guard on; Stop turns patrol off; Restore also turns it off.
    """
    global update_guard_enabled

    while not shutdown_requested.is_set():
        interval = max(2, min(300, int(update_guard_interval_seconds)))
        shutdown_requested.wait(interval)
        if shutdown_requested.is_set() or not update_guard_enabled:
            continue

        try:
            updates = check_update_controls()

            bad_services = [
                svc["name"]
                for svc in updates.get("services", [])
                if svc.get("exists") and not update_service_required_ok(svc)
            ]

            bad_optional = [
                svc["name"]
                for svc in updates.get("optional_services", [])
                if svc.get("exists") and not update_service_optional_ok(svc)
            ]

            bad_policy = not updates.get("policy_disabled")

            if bad_services or bad_optional or bad_policy:
                problems = bad_services + bad_optional
                if bad_policy:
                    problems.append("NoAutoUpdate policy")

                set_current_action("Update Guard")
                add_log("Guard detected re-enabled update control: " + ", ".join(problems))

                results = disable_update_controls()
                log_update_results("Guard", results)

                if not ignore_defender:
                    defender = check_defender()
                    if not defender.get("ok"):
                        add_log("Guard detected Defender check; repair started")
                        defender_results = fix_defender_logic(check_admin=False)
                        log_defender_results("Guard Defender", defender_results)

                add_log("Guard pass completed")
                set_current_action("Idle")

        except Exception as exc:
            add_log(f"Guard error: {exc}")
            set_current_action("Idle")


# -------------------------
# API FOR HTML
# -------------------------

class API:
    def get_status(self):
        return check_status()

    def get_activity_log(self):
        return get_activity_snapshot()

    def open_status_window(self):
        global status_window

        try:
            if status_window is not None:
                status_window.show()
                apply_native_window_icons()
                add_log("Diagnostics window opened")
                return {"ok": True, "opened": True}
        except Exception as exc:
            add_log(f"Diagnostics window failed: {exc}")
            return {"ok": False, "error": str(exc)}

        return {"ok": False, "error": "Diagnostics window is not available"}

    def run_all(self):
        global update_guard_enabled

        set_current_action("Start")
        add_log("Start requested")

        try:
            if not require_admin():
                add_log("Start stopped: admin required")
                return check_status()

            update_results = disable_update_controls()
            log_update_results("Start", update_results)

            update_guard_enabled = True
            add_log("Update guard enabled")

            defender_results = fix_defender_logic(check_admin=False)
            log_defender_results("Start Defender", defender_results)

            status = check_status()
            add_log("Start completed")
            return status
        finally:
            set_current_action("Idle")

    def stop_watchdog(self):
        """Stop the update patrol loop without restoring Windows Update settings."""
        global update_guard_enabled

        set_current_action("Stop")
        add_log("Stop requested")

        try:
            if update_guard_enabled:
                update_guard_enabled = False
                add_log("Update guard disabled")
            else:
                add_log("Update guard already off")

            add_log("Stop completed")
            return check_status()
        finally:
            set_current_action("Idle")

    def fix_defender(self):
        set_current_action("Fix Defender")
        add_log("Fix Defender started")

        try:
            results = fix_defender_logic(check_admin=True)
            if results.get("admin") is False:
                add_log("Fix Defender stopped: admin required")
            else:
                log_defender_results("Fix Defender", results)
                add_log("Fix Defender completed")
            return check_status()
        finally:
            set_current_action("Idle")

    def restore_all(self):
        global update_guard_enabled

        set_current_action("Restore")
        add_log("Restore started")

        try:
            if not require_admin():
                add_log("Restore stopped: admin required")
                return check_status()

            update_guard_enabled = False
            add_log("Update guard disabled")

            results = restore_update_controls()
            log_update_results("Restore", results)
            add_log("Restore completed")
            return check_status()
        finally:
            set_current_action("Idle")

    def toggle_ignore(self, value):
        global ignore_defender
        ignore_defender = bool(value)
        add_log(f"Ignore Defender set to {ignore_defender}")
        return check_status()

    def set_guard_interval(self, value):
        global update_guard_interval_seconds

        try:
            interval = int(value)
        except Exception:
            interval = update_guard_interval_seconds

        interval = max(2, min(300, interval))
        update_guard_interval_seconds = interval
        add_log(f"Watch interval set to {interval} second(s)")
        return check_status()


# -------------------------
# TRAY ICON
# -------------------------

def on_status_closed():
    global status_window
    status_window = None
    add_log("Diagnostics window closed")


def on_status_closing():
    """Hide the diagnostics window without exiting the main watchdog."""
    if status_window is not None and not shutdown_requested.is_set():
        try:
            status_window.hide()
            add_log("Diagnostics window hidden")
            return False
        except Exception:
            pass
    return True


def create_icon(color):
    img = Image.new("RGB", (16, 16), (0, 0, 0))
    d = ImageDraw.Draw(img)
    d.ellipse((2, 2, 14, 14), fill=color)
    return img


def update_tray():
    global icon

    while not shutdown_requested.is_set():
        try:
            status = check_status()
            if icon is not None:
                if status["services"] and (status["defender"] or status["defender"] is None):
                    icon.icon = create_icon("green")
                else:
                    icon.icon = create_icon("red")
        except Exception:
            if icon is not None:
                icon.icon = create_icon("gray")

        shutdown_requested.wait(10)


def open_panel(icon_obj=None, item=None):
    """Show the tiny main window from the system tray."""
    try:
        if window is not None:
            window.show()
            apply_native_window_icons()
            add_log("Main window opened from tray")
            return

        if webview.windows:
            webview.windows[0].show()
            apply_native_window_icons()
            add_log("Main window opened from tray")
    except Exception as exc:
        add_log(f"Open from tray failed: {exc}")


def exit_app(icon_obj=None, item=None):
    """Fully close the tray icon, webview window, and Python process."""
    shutdown_requested.set()

    tray_icon = icon_obj or icon
    if tray_icon is not None:
        try:
            tray_icon.stop()
        except Exception:
            pass

    for item in list(webview.windows):
        try:
            item.destroy()
        except Exception:
            pass

    # pystray/webview can leave platform-specific message loops alive.
    # This app is a local admin utility, so a hard exit is preferable to
    # leaving a hidden watchdog process behind after the user exits.
    os._exit(0)


def on_closing():
    """Minimize the tiny main window to the system tray when X is clicked."""
    if shutdown_requested.is_set():
        return True

    try:
        if window is not None:
            window.hide()
        add_log("Main window hidden to system tray")
        return False
    except Exception as exc:
        add_log(f"Hide to tray failed: {exc}")
        return False


def run_tray():
    global icon

    icon = pystray.Icon(
        "Watchdog",
        create_icon("gray"),
        APP_NAME,
        menu=pystray.Menu(
            pystray.MenuItem("Open", open_panel),
            pystray.MenuItem("Exit", exit_app),
        ),
    )

    threading.Thread(target=update_tray, daemon=True).start()
    icon.run()


# -------------------------
# MAIN
# -------------------------

def main():
    global window, status_window, api_instance

    set_windows_app_id()
    add_log(f"{APP_NAME} v{APP_VERSION} started")
    api_instance = API()

    threading.Thread(target=run_tray, daemon=True).start()
    threading.Thread(target=update_guard_loop, daemon=True).start()

    window = webview.create_window(
        APP_NAME,
        resource_path("ui.html"),
        js_api=api_instance,
        width=202,
        height=128,
        hidden=False,
        resizable=False,
    )

    status_window = webview.create_window(
        "Watchdog Status",
        resource_path("status.html"),
        js_api=api_instance,
        width=680,
        height=520,
        hidden=True,
        resizable=True,
    )

    window.events.loaded += apply_native_window_icons
    status_window.events.loaded += apply_native_window_icons
    window.events.closing += on_closing
    status_window.events.closing += on_status_closing
    webview.start()


if __name__ == "__main__":
    main()
