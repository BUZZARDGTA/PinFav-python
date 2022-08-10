import sys
import win32con
import win32gui
import win32process
import psutil
import colorama
from colorama import Fore
from msvcrt import getch

colorama.init(autoreset=True)

def get_name_and_pid(pid, name):
    for process in psutil.process_iter():
        if pid:
            if pid == process.pid:
                name = process.name()
                return pid, name

        elif name:
            if name.lower() == process.name().lower():
                name = process.name()
                pid = process.pid
                for child_process in psutil.Process(pid=process.pid).parents():
                    if not child_process.name().lower() == process.name().lower():
                        break
                    name = child_process.name()
                    pid = child_process.pid
                return pid, name

    return pid, name

def enum_window_callback(hwnd, pid):
    _, current_pid = win32process.GetWindowThreadProcessId(hwnd) # _ = tid
    if pid == current_pid:
        if win32gui.IsWindowVisible(hwnd):
            hwnd_windows.append(hwnd)

def choice(options):
    def choice_error():
        print(f"\n{Fore.RED}ERROR: Bad option. Options are: {options}")

    while True:
        print(user_action_text, end="")
        try:
            output = getch().decode().upper()
        except UnicodeDecodeError:
            choice_error()
            continue
        if output in options:
            print(f"{Fore.YELLOW}{output}")
            return output
        else:
            choice_error()
            continue

pid = name = None

try:
    pid = int(sys.argv[1])
except ValueError:
    name = str(sys.argv[1])
except IndexError:
    print(f"{Fore.RED}At least one argument is requiered!")
    exit(1)

if name:
    if not name.lower().endswith(".exe"):
        name = f"{name}.exe"

pid, name = get_name_and_pid(pid, name)

if name and not pid:
    print(f"{Fore.RED}No process NAME found for process: {name}!")
    exit(1)
if not pid:
    print(f"{Fore.RED}No process PID found for name: {name}!")
    exit(1)
if not name:
    print(f"{Fore.RED}No process NAME found for PID: {pid}!")
    exit(1)

hwnd_windows = []
win32gui.EnumWindows(enum_window_callback, pid)
if not hwnd_windows:
    print(f"{Fore.RED}For some reasons, HWND have not been found from the window {Fore.GREEN}{name}{Fore.CYAN} ({Fore.GREEN}{pid}{Fore.CYAN})")
    print(f"{Fore.CYAN}You are probably getting this error message for 2 known reasons:")
    print(f"{Fore.CYAN}- Your program is minimized to your system tray, so the script ignores it.")
    print(f"{Fore.CYAN}- Your program doesn't have a GUI (Graphical User Interface), so the script ignores it.")
    exit(1)

user_action_text = f"{Fore.CYAN}Do you want to ({Fore.YELLOW}P{Fore.CYAN})in, ({Fore.YELLOW}U{Fore.CYAN})npin the window {Fore.GREEN}{name}{Fore.CYAN} ({Fore.GREEN}{pid}{Fore.CYAN}) or ({Fore.YELLOW}C{Fore.CYAN})ancel? [{Fore.YELLOW}P{Fore.CYAN}, {Fore.YELLOW}U{Fore.CYAN}, {Fore.YELLOW}C{Fore.CYAN}]: "
user_action = choice(["P", "U", "C"])
if user_action == "C":
    exit(0)

if not psutil.pid_exists(pid):
    print(f"{Fore.RED}Window not found, have you actually closed it?")
    exit(1)

win32gui.ShowWindow(hwnd_windows[0], win32con.SW_SHOW)
if user_action == "P":
    state = win32con.HWND_TOPMOST
    action = "PINNED"
else:
    state = win32con.HWND_NOTOPMOST
    action = "UNPINNED"
flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
win32gui.SetWindowPos(hwnd_windows[0], state, 0,0,0,0, flags)
win32gui.SetForegroundWindow(hwnd_windows[0])

print(f"{Fore.CYAN}Successfully [{Fore.MAGENTA}{action}{Fore.CYAN}] the window.")