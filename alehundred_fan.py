# -*- coding: utf-8 -*-
"""
alehundred_fan.py
HP Victus fan control daemon with system tray.
Hysteresis: above HIGH -> fan max, below LOW -> auto (BIOS).
Based on OmenHwCtl WMI calls (github.com/GeographicCone/OmenHwCtl).
"""

import subprocess
import ctypes
import hashlib
import time
import logging
import threading
import math
import sys
import os
import json
from pathlib import Path
from datetime import datetime

import pystray
from pystray import MenuItem as item
from PIL import Image, ImageDraw

# -- CONFIG --------------------------------------------------------------------

CONFIG_FILE = Path(__file__).parent / "alehundred_fan.json"
LOG_FILE = Path(__file__).parent / "alehundred_fan.log"

DEFAULT_CONFIG = {
    "temp_high": 60,
    "temp_low": 50,
    "check_interval": 5,
}

# -- WMI CONSTANTS (OmenHwCtl) -------------------------------------------------

WMI_COMMAND = 0x20008
WMI_SIGN = [0x53, 0x45, 0x43, 0x55]  # "SECU"
CMD_MAX_FAN = 0x27
CMD_FAN_LEVEL = 0x2D

# -- STATE ---------------------------------------------------------------------

state = {
    "status": "Starting...",
    "temp": 0.0,
    "fan0": 0,
    "fan1": 0,
    "fan_is_max": False,
    "last_check": "Never",
    "transitions": 0,
    "running": True,
    "error": None,
}

config = {}

# -- LOGGING -------------------------------------------------------------------

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-7s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# -- CONFIG I/O ----------------------------------------------------------------

def load_config():
    global config
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r") as f:
                config = json.load(f)
            for k, v in DEFAULT_CONFIG.items():
                if k not in config:
                    config[k] = v
            return
        except Exception:
            pass
    config = dict(DEFAULT_CONFIG)
    save_config()


def save_config():
    with open(CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)

# -- TRAY ICON (FAN) ----------------------------------------------------------

def draw_fan_icon(color="#00CC66", bg=None):
    """Draw a fan icon with 3 curved blades using rotated ellipses."""
    size = 64
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    cx, cy = size // 2, size // 2

    # Background circle (housing)
    draw.ellipse([2, 2, size - 2, size - 2], fill="#1a1a2e", outline="#444466", width=2)

    # Draw 3 fan blades as rotated ellipses
    num_blades = 3
    blade_len = 20
    blade_width = 12

    for i in range(num_blades):
        angle_deg = i * 120 + 30  # offset so it looks dynamic
        # Create a blade on its own image
        blade_img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
        blade_draw = ImageDraw.Draw(blade_img)
        # Draw ellipse centered above hub
        bx = cx - blade_width // 2
        by = cy - blade_len - 4
        blade_draw.ellipse(
            [bx, by, bx + blade_width, by + blade_len],
            fill=color, outline="#FFFFFF", width=1
        )
        # Rotate around center
        blade_img = blade_img.rotate(-angle_deg, center=(cx, cy), resample=Image.BICUBIC)
        img = Image.alpha_composite(img, blade_img)

    # Redraw hub on top
    draw2 = ImageDraw.Draw(img)
    r_hub = 6
    draw2.ellipse(
        [cx - r_hub, cy - r_hub, cx + r_hub, cy + r_hub],
        fill="#FFFFFF", outline=color, width=2
    )
    # Center dot
    draw2.ellipse(
        [cx - 2, cy - 2, cx + 2, cy + 2],
        fill=color
    )

    return img


def update_tray(tray):
    if state.get("error"):
        tray.icon = draw_fan_icon("#CC3333")
    elif state["fan_is_max"]:
        tray.icon = draw_fan_icon("#FF6600")
    else:
        tray.icon = draw_fan_icon("#00CC66")
    tray.title = build_tooltip()


def build_tooltip():
    temp = state["temp"]
    mode = "MAX" if state["fan_is_max"] else "AUTO"
    fans = "{}/{}".format(state["fan0"], state["fan1"])
    return "Alehundred Fan -- {:.0f}C {} | Fans {} RPM | H: {}/{}".format(
        temp, mode, fans, config.get("temp_low", "?"), config.get("temp_high", "?")
    )

# -- WMI CALLS ----------------------------------------------------------------

def ps(command, timeout=10):
    try:
        r = subprocess.run(
            ["powershell", "-NoProfile", "-Command", command],
            capture_output=True, text=True, timeout=timeout,
            creationflags=subprocess.CREATE_NO_WINDOW,
        )
        return r.stdout.strip()
    except Exception:
        return ""


def call_wmi(command_type, data=None, output_size=0):
    if data is None:
        data = []

    sign_arr = ','.join(str(b) for b in WMI_SIGN)
    data_arr = ','.join(str(b) for b in data) if data else ''
    method = "hpqBIOSInt{}".format(output_size)

    script = r"""
try {{
    [byte[]]$sign = @({sign})
    [uint32]$cmd = {cmd}
    [uint32]$cmdType = {cmd_type}
    [uint32]$size = {size}
    {data_line}

    $dataIn = New-CimInstance -ClassName hpqBDataIn -Namespace root/WMI -ClientOnly -Property @{{
        Sign        = $sign
        Command     = $cmd
        CommandType = $cmdType
        Size        = $size
        hpqBData    = $hpqBData
    }}

    $inst = Get-CimInstance -Namespace root/WMI -ClassName hpqBIntM | Select-Object -First 1
    if (-not $inst) {{ Write-Output "ERR"; exit }}

    $result = Invoke-CimMethod -InputObject $inst -MethodName {method} -Arguments @{{
        InData = [CimInstance]$dataIn
    }}

    $out = $result.OutData
    if ($out) {{
        $s = [System.Text.Encoding]::ASCII.GetString($out.CimInstanceProperties['Sign'].Value[0..3])
        $rc = $out.CimInstanceProperties['rwReturnCode'].Value
        $d = $out.CimInstanceProperties['Data'].Value
        if ($d) {{
            $hex = ($d[0..[Math]::Min(7, $d.Length-1)] | ForEach-Object {{ $_.ToString('X2') }}) -join ' '
            Write-Output "$s|$rc|$hex"
        }} else {{
            Write-Output "$s|$rc|"
        }}
    }} else {{
        Write-Output "OK|0|"
    }}
}} catch {{
    Write-Output "ERR|9|$($_.Exception.Message)"
}}
""".format(
        sign=sign_arr,
        cmd=WMI_COMMAND,
        cmd_type=command_type,
        size=len(data),
        data_line='[byte[]]$hpqBData = @({})'.format(data_arr) if data else '[byte[]]$hpqBData = @()',
        method=method,
    )
    return ps(script, timeout=10)


def read_temperature():
    out = ps(
        r"(Get-CimInstance -Namespace root\WMI -ClassName MSAcpi_ThermalZoneTemperature"
        r" -ErrorAction SilentlyContinue).CurrentTemperature"
    )
    if out and out.isdigit():
        return round((int(out) / 10.0) - 273.15, 1)
    return None


def fan_max_on():
    result = call_wmi(CMD_MAX_FAN, [0x01], output_size=0)
    return result and 'PASS' in result


def fan_max_off():
    result = call_wmi(CMD_MAX_FAN, [0x00], output_size=0)
    return result and 'PASS' in result


def get_fan_speed():
    result = call_wmi(CMD_FAN_LEVEL, [], output_size=128)
    if result and 'PASS' in result:
        parts = result.split('|')
        if len(parts) >= 3 and parts[2].strip():
            hx = parts[2].strip().split(' ')
            if len(hx) >= 2:
                return int(hx[0], 16) * 100, int(hx[1], 16) * 100
    return 0, 0

# -- MAIN LOOP ----------------------------------------------------------------

def fan_loop(tray):
    log.info("=" * 60)
    log.info("Alehundred Fan started")
    log.info("  Hysteresis: >%d C -> MAX | <%d C -> AUTO",
             config["temp_high"], config["temp_low"])
    log.info("  Interval: %ds", config["check_interval"])
    log.info("=" * 60)

    # Start in MAX to protect during boot
    fan_max_on()
    state["fan_is_max"] = True
    state["status"] = "Watching..."
    log.info("  Initial mode: MAX (boot protection)")

    while state["running"]:
        time.sleep(config["check_interval"])
        if not state["running"]:
            break

        temp = read_temperature()
        fan0, fan1 = get_fan_speed()

        if temp is None:
            state["error"] = "Cannot read temperature"
            state["status"] = "ERROR: temp read failed"
            update_tray(tray)
            log.error("Temperature read failed")
            continue

        state["temp"] = temp
        state["fan0"] = fan0
        state["fan1"] = fan1
        state["error"] = None
        state["last_check"] = datetime.now().strftime("%H:%M:%S")

        high = config["temp_high"]
        low = config["temp_low"]

        if temp > high and not state["fan_is_max"]:
            ok = fan_max_on()
            if ok:
                state["fan_is_max"] = True
                state["transitions"] += 1
                state["status"] = "{:.0f}C > {}C -> MAX".format(temp, high)
                log.info("FAN MAX ON  %.1f C > %d C  | Fans: %d/%d RPM",
                         temp, high, fan0, fan1)
            else:
                state["error"] = "WMI call failed"
                state["status"] = "ERROR: fan max on failed"
                log.error("Fan max on failed at %.1f C", temp)

        elif temp < low and state["fan_is_max"]:
            ok = fan_max_off()
            if ok:
                state["fan_is_max"] = False
                state["transitions"] += 1
                state["status"] = "{:.0f}C < {}C -> AUTO".format(temp, low)
                log.info("FAN AUTO  %.1f C < %d C  | Fans: %d/%d RPM",
                         temp, low, fan0, fan1)
            else:
                state["error"] = "WMI call failed"
                state["status"] = "ERROR: fan auto failed"
                log.error("Fan auto failed at %.1f C", temp)

        else:
            mode = "MAX" if state["fan_is_max"] else "AUTO"
            state["status"] = "{:.0f}C | {} | {}/{}".format(temp, mode, fan0, fan1)
            log.info("%.1f C | %s | Fans: %d/%d RPM", temp, mode, fan0, fan1)
            # Re-send max command every cycle to prevent OMEN/BIOS from overriding
            if state["fan_is_max"]:
                fan_max_on()

        update_tray(tray)

    # Cleanup
    log.info("Shutting down, setting fan to AUTO...")
    fan_max_off()
    log.info("Stopped.")

# -- TRAY MENU ACTIONS ---------------------------------------------------------

def show_status(icon, _):
    mode = "MAX" if state["fan_is_max"] else "AUTO"
    lines = [
        "Temp: {:.1f} C  |  Mode: {}".format(state["temp"], mode),
        "Fans: {} / {} RPM".format(state["fan0"], state["fan1"]),
        "Hysteresis: <{} AUTO | >{} MAX".format(config["temp_low"], config["temp_high"]),
        "Last check: {}".format(state["last_check"]),
        "Transitions: {}".format(state["transitions"]),
        state["status"],
    ]
    if state.get("error"):
        lines.append("ERROR: {}".format(state["error"]))
    icon.notify("\n".join(lines), "Alehundred Fan")


def open_log(icon, _):
    subprocess.Popen(["notepad.exe", str(LOG_FILE)])


def set_thresholds(icon, _):
    """Open a small tkinter dialog to set LOW/HIGH thresholds."""
    def _dialog():
        import tkinter as tk

        root = tk.Tk()
        root.title("Alehundred Fan - Thresholds")
        root.resizable(False, False)
        root.attributes("-topmost", True)

        # Center on screen
        w, h = 320, 200
        sx = root.winfo_screenwidth() // 2 - w // 2
        sy = root.winfo_screenheight() // 2 - h // 2
        root.geometry("{}x{}+{}+{}".format(w, h, sx, sy))
        root.configure(bg="#1a1a2e")

        tk.Label(
            root, text="Fan Hysteresis", font=("Segoe UI", 14, "bold"),
            bg="#1a1a2e", fg="#FFFFFF"
        ).pack(pady=(15, 10))

        frame = tk.Frame(root, bg="#1a1a2e")
        frame.pack(pady=5)

        tk.Label(
            frame, text="LOW (auto):", font=("Segoe UI", 10),
            bg="#1a1a2e", fg="#00CC66"
        ).grid(row=0, column=0, padx=10, pady=5, sticky="e")

        low_var = tk.StringVar(value=str(config["temp_low"]))
        low_entry = tk.Entry(frame, textvariable=low_var, width=6,
                             font=("Segoe UI", 11), justify="center")
        low_entry.grid(row=0, column=1, padx=5)

        tk.Label(
            frame, text="C", font=("Segoe UI", 10),
            bg="#1a1a2e", fg="#AAAAAA"
        ).grid(row=0, column=2)

        tk.Label(
            frame, text="HIGH (max):", font=("Segoe UI", 10),
            bg="#1a1a2e", fg="#FF6600"
        ).grid(row=1, column=0, padx=10, pady=5, sticky="e")

        high_var = tk.StringVar(value=str(config["temp_high"]))
        high_entry = tk.Entry(frame, textvariable=high_var, width=6,
                              font=("Segoe UI", 11), justify="center")
        high_entry.grid(row=1, column=1, padx=5)

        tk.Label(
            frame, text="C", font=("Segoe UI", 10),
            bg="#1a1a2e", fg="#AAAAAA"
        ).grid(row=1, column=2)

        status_label = tk.Label(
            root, text="", font=("Segoe UI", 9),
            bg="#1a1a2e", fg="#CC3333"
        )
        status_label.pack()

        def apply():
            try:
                lo = int(low_var.get().strip())
                hi = int(high_var.get().strip())
            except ValueError:
                status_label.config(text="Numbers only")
                return
            if lo >= hi:
                status_label.config(text="LOW must be < HIGH")
                return
            if lo < 30 or hi > 95:
                status_label.config(text="Range: 30-95 C")
                return
            config["temp_low"] = lo
            config["temp_high"] = hi
            save_config()
            log.info("Thresholds changed: LOW=%d HIGH=%d", lo, hi)
            root.destroy()

        tk.Button(
            root, text="Apply", command=apply,
            font=("Segoe UI", 10, "bold"), bg="#00CC66", fg="#000000",
            activebackground="#00AA55", width=10, relief="flat", cursor="hand2"
        ).pack(pady=(10, 5))

        root.mainloop()

    t = threading.Thread(target=_dialog, daemon=True)
    t.start()


def install_task(icon, _):
    """Register as a scheduled task that runs at logon with admin privileges."""
    script_path = os.path.abspath(__file__)
    # Use pythonw.exe (no console window) instead of python.exe
    python_path = sys.executable.replace("python.exe", "pythonw.exe")

    xml = r"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Alehundred Fan Control - HP Victus thermal management</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <LogonType>InteractiveToken</LogonType>
      <RunLevel>HighestAvailable</RunLevel>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>true</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT0S</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>"{python}"</Command>
      <Arguments>"{script}"</Arguments>
      <WorkingDirectory>{workdir}</WorkingDirectory>
    </Exec>
  </Actions>
</Task>""".format(
        python=python_path,
        script=script_path,
        workdir=os.path.dirname(script_path),
    )

    xml_path = Path(os.environ.get("TEMP", ".")) / "alehundred_fan_task.xml"
    with open(xml_path, "w", encoding="utf-16") as f:
        f.write(xml)

    result = subprocess.run(
        ["schtasks", "/Create", "/TN", "AlehundredFan", "/XML", str(xml_path), "/F"],
        capture_output=True, text=True
    )

    if result.returncode == 0:
        log.info("Task Scheduler: installed successfully")
        icon.notify("Task installed!\nWill start at logon with admin.", "Alehundred Fan")
    else:
        log.error("Task Scheduler install failed: %s", result.stderr)
        icon.notify("Install failed: {}".format(result.stderr[:200]), "Alehundred Fan")

    try:
        xml_path.unlink()
    except Exception:
        pass


def uninstall_task(icon, _):
    result = subprocess.run(
        ["schtasks", "/Delete", "/TN", "AlehundredFan", "/F"],
        capture_output=True, text=True
    )
    if result.returncode == 0:
        log.info("Task Scheduler: uninstalled")
        icon.notify("Task removed.", "Alehundred Fan")
    else:
        log.error("Task Scheduler uninstall failed: %s", result.stderr)
        icon.notify("Uninstall failed: {}".format(result.stderr[:200]), "Alehundred Fan")


def quit_app(icon, _):
    state["running"] = False
    icon.stop()

# -- ADMIN CHECK ---------------------------------------------------------------

def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except Exception:
        return False


def relaunch_as_admin():
    """Relaunch this script as admin via UAC prompt."""
    script = os.path.abspath(__file__)
    pythonw = sys.executable.replace("python.exe", "pythonw.exe")
    params = '"{}"'.format(script)
    try:
        ctypes.windll.shell32.ShellExecuteW(
            None, "runas", pythonw, params, None, 1
        )
    except Exception:
        pass
    sys.exit(0)

# -- MAIN ----------------------------------------------------------------------

def main():
    if not is_admin():
        relaunch_as_admin()

    load_config()

    icon_img = draw_fan_icon("#00CC66")

    menu = pystray.Menu(
        item("Alehundred Fan", lambda i, it: None, enabled=False),
        pystray.Menu.SEPARATOR,
        item("Show status", show_status),
        item("Set thresholds", set_thresholds),
        pystray.Menu.SEPARATOR,
        item("Install at startup", install_task),
        item("Remove from startup", uninstall_task),
        pystray.Menu.SEPARATOR,
        item("Open log", open_log),
        item("Quit", quit_app),
    )

    tray = pystray.Icon(
        name="AlehundredFan",
        icon=icon_img,
        title="Alehundred Fan -- Starting...",
        menu=menu,
    )

    t = threading.Thread(target=fan_loop, args=(tray,), daemon=True)
    t.start()

    tray.run()


if __name__ == "__main__":
    main()