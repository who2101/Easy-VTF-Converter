import win32gui as w
import comtypes.client
import pygetwindow as gw
import os



def _normalise_text(controlText):
    return controlText.lower().replace('&', '')


def _window_enumeration_handler(hwnd, result_list):
    result_list.append((hwnd, w.GetWindowText(hwnd), w.GetClassName(hwnd)))


def find_child_windows(current_hwnd, wanted_class=None):
    results = []
    children = []

    try:
        w.EnumChildWindows(current_hwnd, _window_enumeration_handler, children)
    except w.error:
        return

    for child_hwnd, _, window_class in children:
        if wanted_class and not window_class == wanted_class:
            continue

        results.append(child_hwnd)

    return results


def window_iterator(hwnd, output):
    if w.IsWindowVisible(hwnd) and w.GetClassName(hwnd) == "CabinetWClass":
        output.append(hwnd)


def get_explorer_window_paths():
    windows = gw.getAllWindows()

    paths = []

    for window in windows:
        if w.GetClassName(window._hWnd) == "CabinetWClass":
            hwnd = window._hWnd

            current_path = get_explorer_window_path(hwnd)

            if current_path:
                paths.append(current_path)
    
    return paths
            

def get_explorer_window_path(MyHwnd):
    try:
        shell_app = comtypes.client.CreateObject("Shell.Application")
        windows = shell_app.Windows()
        for i in range(windows.Count):
            explorer_window = windows.Item(i)
 
            if explorer_window is None:
                continue
            if explorer_window.hwnd == MyHwnd:
                path = os.path.basename(explorer_window.FullName)
 
                if path.lower() == "explorer.exe":
                    explore_path = explorer_window.Document.Folder.Items().Item().Path
                    return explore_path
    finally:
        shell_app = None