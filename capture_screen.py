"""全屏截图（含 Windows 任务栏系统时间，作为核查时点证据）。

只截「浏览器当前所在的那一块显示器」，不截整个虚拟桌面。
多显示器时 ImageGrab.grab(all_screens=True) 会把两块屏加中间空隙一起拍成
一张巨图（例：5120x2000，有效内容只占三分之一），塞进 Word 缩到页宽后看不清。

用法:
  capture_screen.py <output.jpg>              # 截前台窗口所在的显示器
  capture_screen.py <output.jpg> --hwnd 1234  # 截该窗口所在的显示器（run.ps1 传浏览器句柄）
  capture_screen.py <output.jpg> --monitor 1  # 强制截 1 号显示器（主屏）
  capture_screen.py <output.jpg> --all        # 旧行为：整个虚拟桌面
"""

from pathlib import Path
import ctypes
from ctypes import wintypes
import sys
import time

from PIL import ImageGrab

# 必须在任何窗口/显示器 API 之前设置，否则拿到的是被 DPI 缩放过的逻辑坐标，
# 与 ImageGrab 的物理像素对不上，截出来会错位。
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)  # PER_MONITOR_DPI_AWARE
except Exception:
    ctypes.windll.user32.SetProcessDPIAware()

user32 = ctypes.windll.user32

MONITOR_DEFAULTTONEAREST = 2
MONITORINFOF_PRIMARY = 1
SM_XVIRTUALSCREEN, SM_YVIRTUALSCREEN = 76, 77


class MONITORINFOEXW(ctypes.Structure):
    _fields_ = [("cbSize", wintypes.DWORD),
                ("rcMonitor", wintypes.RECT),
                ("rcWork", wintypes.RECT),
                ("dwFlags", wintypes.DWORD),
                ("szDevice", wintypes.WCHAR * 32)]


MONITORENUMPROC = ctypes.WINFUNCTYPE(wintypes.BOOL, wintypes.HMONITOR,
                                     wintypes.HDC, ctypes.POINTER(wintypes.RECT),
                                     wintypes.LPARAM)


def _info(hmon) -> MONITORINFOEXW:
    mi = MONITORINFOEXW()
    mi.cbSize = ctypes.sizeof(MONITORINFOEXW)
    user32.GetMonitorInfoW(hmon, ctypes.byref(mi))
    return mi


def list_monitors() -> list:
    found = []

    def collect(hmon, hdc, lprc, lparam):
        found.append(_info(hmon))
        return True

    user32.EnumDisplayMonitors(0, None, MONITORENUMPROC(collect), 0)
    # 主屏排在前面，这样 --monitor 1 就是主屏
    found.sort(key=lambda m: 0 if m.dwFlags & MONITORINFOF_PRIMARY else 1)
    return found


def monitor_of_window(hwnd=None) -> MONITORINFOEXW:
    """指定窗口（默认前台窗口）所在的显示器；窗口无效时退回主屏。"""
    if not hwnd:
        hwnd = user32.GetForegroundWindow()
    if not hwnd or not user32.IsWindow(hwnd):
        return list_monitors()[0]
    return _info(user32.MonitorFromWindow(hwnd, MONITOR_DEFAULTTONEAREST))


def main() -> int:
    args = sys.argv[1:]
    if not args:
        print("Usage: capture_screen.py <output.jpg> [--monitor N | --all]",
              file=sys.stderr)
        return 2

    output = Path(args[0])
    grab_all = "--all" in args
    monitor_index = None
    if "--monitor" in args:
        try:
            monitor_index = int(args[args.index("--monitor") + 1])
        except (IndexError, ValueError):
            print("--monitor 后面要跟显示器编号，例如 --monitor 1", file=sys.stderr)
            return 2

    hwnd = None
    if "--hwnd" in args:
        try:
            hwnd = int(args[args.index("--hwnd") + 1])
        except (IndexError, ValueError):
            hwnd = None

    output.parent.mkdir(parents=True, exist_ok=True)
    time.sleep(0.5)

    if grab_all:
        image = ImageGrab.grab(all_screens=True)
        where = "整个虚拟桌面"
    else:
        monitors = list_monitors()
        if monitor_index is not None:
            if not 1 <= monitor_index <= len(monitors):
                print(f"只有 {len(monitors)} 块显示器，没有 {monitor_index} 号", file=sys.stderr)
                return 2
            mi = monitors[monitor_index - 1]
        else:
            mi = monitor_of_window(hwnd)

        r = mi.rcMonitor
        # ImageGrab 的 bbox 以虚拟桌面左上角为原点；副屏在主屏左侧/上方时该原点为负
        ox = user32.GetSystemMetrics(SM_XVIRTUALSCREEN)
        oy = user32.GetSystemMetrics(SM_YVIRTUALSCREEN)
        bbox = (r.left - ox, r.top - oy, r.right - ox, r.bottom - oy)
        image = ImageGrab.grab(bbox=bbox, all_screens=True)
        where = f"{mi.szDevice} ({r.right - r.left}x{r.bottom - r.top})"

    if image.mode != "RGB":
        image = image.convert("RGB")

    # 截图是「大片纯色 + 细小文字」，PNG 无损且比高质量 JPEG 还小，优先用 PNG。
    # 仍存 JPEG 时必须关掉色度抽样(subsampling=0)，否则红底白字的页面边缘会糊成一片。
    if output.suffix.lower() in (".jpg", ".jpeg"):
        image.save(output, quality=95, subsampling=0)
    else:
        image.save(output, optimize=True)

    size_kb = output.stat().st_size / 1024
    print(f"{output}  [{where}  {image.width}x{image.height}  {size_kb:.0f}KB]")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
