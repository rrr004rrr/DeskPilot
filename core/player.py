"""
Player Module
Executes recorded scripts using pyautogui.
"""
import time
import threading
import pyautogui

pyautogui.FAILSAFE = False  # disable corner failsafe for automation use


MAX_CALL_DEPTH = 10  # 防止循環呼叫造成無限遞迴


class Player:
    def __init__(self):
        self._thread: threading.Thread | None = None
        self._stop_event = threading.Event()
        self._playing = False
        self._on_finished = None  # callback()
        self._on_progress = None  # callback(current_iter, total_iter)
        self._script_loader = None  # fn(name: str) -> list[dict]

    def set_script_loader(self, fn):
        """設定子腳本載入器，fn(name) -> list[dict]"""
        self._script_loader = fn

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------
    def play(self, events: list[dict], repeat: int = 1, loop: bool = False):
        """
        Play events.
        repeat: number of times (ignored if loop=True)
        loop: repeat indefinitely until stop() is called
        """
        if self._playing:
            return
        self._stop_event.clear()
        self._playing = True
        self._thread = threading.Thread(
            target=self._run,
            args=(events, repeat, loop),
            daemon=True,
        )
        self._thread.start()

    def stop(self):
        self._stop_event.set()
        self._playing = False

    def is_playing(self) -> bool:
        return self._playing

    def set_on_finished(self, fn):
        self._on_finished = fn

    def set_on_progress(self, fn):
        """fn(current_iter: int, total_iter: int | None)"""
        self._on_progress = fn

    # ------------------------------------------------------------------
    # Internal
    # ------------------------------------------------------------------
    def _run(self, events: list[dict], repeat: int, loop: bool):
        iteration = 0
        try:
            while not self._stop_event.is_set():
                iteration += 1
                total = None if loop else repeat
                if self._on_progress:
                    self._on_progress(iteration, total)
                self._execute_events(events)
                if not loop and iteration >= repeat:
                    break
        finally:
            self._playing = False
            if self._on_finished:
                self._on_finished()

    def _execute_events(self, events: list[dict], _depth: int = 0):
        for event in events:
            if self._stop_event.is_set():
                return
            etype = event.get("type")

            if etype == "sleep":
                ms = event.get("ms", 0)
                # sleep in small chunks to allow early stop
                end = time.time() + ms / 1000.0
                while time.time() < end:
                    if self._stop_event.is_set():
                        return
                    time.sleep(0.01)

            elif etype == "mouse_move":
                pyautogui.moveTo(event["x"], event["y"], duration=0)

            elif etype == "click":
                btn = event.get("button", "left")
                pyautogui.mouseDown(event["x"], event["y"], button=btn)

            elif etype == "release":
                btn = event.get("button", "left")
                pyautogui.mouseUp(event["x"], event["y"], button=btn)

            elif etype == "scroll":
                pyautogui.scroll(event.get("dy", 0), x=event["x"], y=event["y"])

            elif etype == "key":
                key_str = event.get("key", "")
                self._press_key(key_str)

            elif etype == "call_script":
                name = event.get("name", "")
                if name and self._script_loader and _depth < MAX_CALL_DEPTH:
                    try:
                        sub_events = self._script_loader(name)
                        self._execute_events(sub_events, _depth + 1)
                    except Exception:
                        pass

    @staticmethod
    def _press_key(key_str: str):
        """Map pynput key strings to pyautogui key names."""
        mapping = {
            "Key.space": "space",
            "Key.enter": "enter",
            "Key.backspace": "backspace",
            "Key.tab": "tab",
            "Key.shift": "shift",
            "Key.shift_r": "shiftright",
            "Key.ctrl_l": "ctrlleft",
            "Key.ctrl_r": "ctrlright",
            "Key.alt_l": "altleft",
            "Key.alt_r": "altright",
            "Key.delete": "delete",
            "Key.home": "home",
            "Key.end": "end",
            "Key.page_up": "pageup",
            "Key.page_down": "pagedown",
            "Key.up": "up",
            "Key.down": "down",
            "Key.left": "left",
            "Key.right": "right",
            "Key.f1": "f1", "Key.f2": "f2", "Key.f3": "f3",
            "Key.f4": "f4", "Key.f5": "f5", "Key.f6": "f6",
            "Key.f7": "f7", "Key.f8": "f8", "Key.f9": "f9",
            "Key.f10": "f10", "Key.f11": "f11", "Key.f12": "f12",
        }
        mapped = mapping.get(key_str)
        if mapped:
            pyautogui.press(mapped)
        elif len(key_str) == 1:
            pyautogui.press(key_str)
