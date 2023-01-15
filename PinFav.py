import sys
import psutil
import colorama
import win32gui
import win32con
import win32process
from colorama import Fore
from msvcrt import getch

colorama.init(autoreset=True)


def get_pid_name__or__parent_pid_from_name():
    global name, pid

    for process in psutil.process_iter():
        if pid:
            if pid == process.pid:
                name = process.name()
                return

        elif name:
            if name.lower() == process.name().lower():
                found_a_parent_process__flag = False
                for parent_process in psutil.Process(pid=process.pid).parents():
                    if not parent_process.name().lower() == process.name().lower():
                        break
                    found_a_parent_process__flag = True
                    name = parent_process.name()
                    pid = parent_process.pid

                if not found_a_parent_process__flag:
                    name = process.name()
                    pid = process.pid
                return

def get_hwnd_by_pid(pid):
    def enum_window_callback(hwnd, pid):
        if pid == win32process.GetWindowThreadProcessId(hwnd)[1]: # [0], [1] = tid, pid
            hwnd_windows.append(hwnd)

    hwnd_windows = []
    win32gui.EnumWindows(enum_window_callback, pid)
    for hwnd in hwnd_windows:
        if win32gui.IsWindowVisible(hwnd):
            return hwnd

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
        if not output in options:
            choice_error()
            continue
        print(f"{Fore.YELLOW}{output}")
        return output

name = None
pid = None
hwnd = None

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

get_pid_name__or__parent_pid_from_name()

if name and not pid:
    print(f"{Fore.RED}No process NAME found for process: {name}!")
    exit(1)
if not pid:
    print(f"{Fore.RED}No process PID found for name: {name}!")
    exit(1)
if not name:
    print(f"{Fore.RED}No process NAME found for PID: {pid}!")
    exit(1)

hwnd = get_hwnd_by_pid(pid)
if not hwnd:
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

win32gui.ShowWindow(hwnd, win32con.SW_SHOW)
if user_action == "P":
    state = win32con.HWND_TOPMOST
    action = "PINNED"
else:
    state = win32con.HWND_NOTOPMOST
    action = "UNPINNED"
flags = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
win32gui.SetWindowPos(hwnd, state, 0,0,0,0, flags)
win32gui.SetForegroundWindow(hwnd)

print(f"{Fore.CYAN}Successfully [{Fore.MAGENTA}{action}{Fore.CYAN}] the window.")