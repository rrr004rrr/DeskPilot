"""
Recorder Module
Records mouse and keyboard events using pynput.
"""
import time
import threading
from pynput import mouse, keyboard


class Recorder:
    def __init__(self):
        self._events: list[dict] = []
        self._recording = False
        self._last_time = 0.0
        self._mouse_listener = None
        self._keyboard_listener = None
        self._lock = threading.Lock()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def start(self):
        if self._recording:
            return
        self._events = []
        self._recording = True
        self._last_time = time.time()

        self._mouse_listener = mouse.Listener(
            on_move=self._on_move,
            on_click=self._on_click,
            on_scroll=self._on_scroll,
        )
        self._keyboard_listener = keyboard.Listener(
            on_press=self._on_key_press,
        )
        self._mouse_listener.start()
        self._keyboard_listener.start()

    def stop(self) -> list[dict]:
        if not self._recording:
            return self._events
        self._recording = False
        if self._mouse_listener:
            self._mouse_listener.stop()
        if self._keyboard_listener:
            self._keyboard_listener.stop()
        return self._events

    def is_recording(self) -> bool:
        return self._recording

    def get_events(self) -> list[dict]:
        return list(self._events)

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------
    def _record_delay(self):
        """Insert a sleep event based on elapsed time since last recorded event."""
        now = time.time()
        ms = int((now - self._last_time) * 1000)
        self._last_time = now
        if ms > 10:
            self._events.append({"type": "sleep", "ms": ms})

    def _on_move(self, x, y):
        pass  # 移動不記錄，位置由點擊事件帶入

    def _on_click(self, x, y, button, pressed):
        if not self._recording:
            return
        with self._lock:
            self._record_delay()
            btn_name = "left" if button == mouse.Button.left else "right"
            if pressed:
                # 先移動到點擊位置，再記錄點擊
                self._events.append({"type": "mouse_move", "x": x, "y": y})
                self._events.append({"type": "click", "button": btn_name, "x": x, "y": y})
            else:
                self._events.append({"type": "release", "button": btn_name, "x": x, "y": y})

    def _on_scroll(self, x, y, dx, dy):
        if not self._recording:
            return
        with self._lock:
            self._record_delay()
            self._events.append({"type": "scroll", "x": x, "y": y, "dx": dx, "dy": dy})

    def _on_key_press(self, key):
        if not self._recording:
            return
        with self._lock:
            self._record_delay()
            try:
                key_str = key.char if hasattr(key, "char") and key.char else str(key)
            except Exception:
                key_str = str(key)
            self._events.append({"type": "key", "key": key_str})
