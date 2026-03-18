"""
Window Manager Module
Handles window listing, selection, position/size control, and locking.
"""
import threading
import time
import win32gui
import win32con
import win32api


class WindowManager:
    def __init__(self):
        self._hwnd = None
        self._locked = False
        self._target = {"x": 0, "y": 0, "w": 800, "h": 600}
        self._always_on_top = False
        self._monitor_thread = None
        self._stop_event = threading.Event()
        self._on_restored = None  # callback

    # ------------------------------------------------------------------
    # Window enumeration
    # ------------------------------------------------------------------
    def get_window_list(self) -> list[dict]:
        """Return list of visible windows with title and hwnd."""
        results = []

        def enum_callback(hwnd, _):
            if win32gui.IsWindowVisible(hwnd):
                title = win32gui.GetWindowText(hwnd)
                if title.strip():
                    results.append({"hwnd": hwnd, "title": title})

        win32gui.EnumWindows(enum_callback, None)
        results.sort(key=lambda w: w["title"].lower())
        return results

    # ------------------------------------------------------------------
    # Selection
    # ------------------------------------------------------------------
    def select_window(self, hwnd: int):
        self._hwnd = hwnd
        rect = self.get_current_rect()
        if rect:
            self._target = rect

    def get_selected_hwnd(self) -> int | None:
        return self._hwnd

    def is_valid(self) -> bool:
        return self._hwnd is not None and win32gui.IsWindow(self._hwnd)

    # ------------------------------------------------------------------
    # Position / size
    # ------------------------------------------------------------------
    def get_current_rect(self) -> dict | None:
        if not self.is_valid():
            return None
        try:
            rect = win32gui.GetWindowRect(self._hwnd)
            x, y = rect[0], rect[1]
            w = rect[2] - rect[0]
            h = rect[3] - rect[1]
            return {"x": x, "y": y, "w": w, "h": h}
        except Exception:
            return None

    def set_target(self, x: int, y: int, w: int, h: int):
        self._target = {"x": x, "y": y, "w": w, "h": h}

    def apply_position(self):
        """Move and resize the selected window to target values."""
        if not self.is_valid():
            return False
        try:
            flags = win32con.SWP_NOZORDER
            if self._always_on_top:
                insert_after = win32con.HWND_TOPMOST
            else:
                insert_after = win32con.HWND_NOTOPMOST
                flags = 0  # allow z-order change

            win32gui.SetWindowPos(
                self._hwnd,
                insert_after,
                self._target["x"],
                self._target["y"],
                self._target["w"],
                self._target["h"],
                win32con.SWP_SHOWWINDOW,
            )
            return True
        except Exception:
            return False

    def set_always_on_top(self, enabled: bool):
        self._always_on_top = enabled
        if self.is_valid():
            insert_after = win32con.HWND_TOPMOST if enabled else win32con.HWND_NOTOPMOST
            try:
                win32gui.SetWindowPos(
                    self._hwnd,
                    insert_after,
                    0, 0, 0, 0,
                    win32con.SWP_NOMOVE | win32con.SWP_NOSIZE,
                )
            except Exception:
                pass

    # ------------------------------------------------------------------
    # Lock / unlock
    # ------------------------------------------------------------------
    def lock(self, interval_ms: int = 200):
        if self._locked:
            return
        self._locked = True
        self._stop_event.clear()
        self.apply_position()
        self._monitor_thread = threading.Thread(
            target=self._monitor_loop,
            args=(interval_ms / 1000.0,),
            daemon=True,
        )
        self._monitor_thread.start()

    def unlock(self):
        self._locked = False
        self._stop_event.set()

    def is_locked(self) -> bool:
        return self._locked

    def set_on_restored_callback(self, fn):
        """Called whenever the window is auto-restored."""
        self._on_restored = fn

    def _monitor_loop(self, interval: float):
        while not self._stop_event.wait(interval):
            if not self.is_valid():
                continue
            rect = self.get_current_rect()
            if rect and rect != self._target:
                self.apply_position()
                if self._on_restored:
                    self._on_restored()

    # ------------------------------------------------------------------
    # Serialisation helpers
    # ------------------------------------------------------------------
    def get_config(self) -> dict:
        return {
            "hwnd": self._hwnd,
            "target": self._target.copy(),
            "always_on_top": self._always_on_top,
        }

    def load_config(self, config: dict):
        self._target = config.get("target", self._target)
        self._always_on_top = config.get("always_on_top", False)
        hwnd = config.get("hwnd")
        if hwnd and win32gui.IsWindow(hwnd):
            self._hwnd = hwnd
