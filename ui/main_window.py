"""
Main Window UI — DeskPilot
"""
from __future__ import annotations

from pathlib import Path

from PyQt6.QtCore import Qt, pyqtSignal, QObject
from PyQt6.QtGui import QColor, QFont, QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QComboBox, QPushButton, QSpinBox, QCheckBox,
    QTableWidget, QTableWidgetItem, QHeaderView, QStatusBar,
    QSplitter, QTabWidget, QFileDialog, QInputDialog, QMessageBox,
    QFrame, QSizePolicy,
)

from core.window_manager import WindowManager
from core.recorder import Recorder
from core.player import Player
from core.script_manager import ScriptManager

# ---------------------------------------------------------------------------
# Palette
# ---------------------------------------------------------------------------
C = {
    "bg":        "#f4f6f9",   # 頁面底色
    "surface":   "#ffffff",   # 卡片/面板
    "border":    "#dde3ec",   # 邊框
    "border2":   "#c8d0dc",   # 加深邊框
    "text":      "#1e2430",   # 主文字
    "muted":     "#6b7a99",   # 次要文字
    "subtle":    "#9aa5bb",   # 提示文字
    "accent":    "#3b7ddd",   # 藍色強調
    "bar_bg":    "#2c3e50",   # Action bar 底色
    "bar_text":  "#ecf0f1",   # Action bar 文字
    "rec":       "#e74c3c",   # 錄製紅
    "rec_hover": "#c0392b",
    "rec_dim":   "#fadbd8",
    "play":      "#27ae60",   # 播放綠
    "play_hover":"#1e8449",
    "play_dim":  "#d5f5e3",
    "stop":      "#7f8c8d",   # 停止灰
    "stop_hover":"#636e72",
    "danger_txt":"#c0392b",   # 刪除紅字
    "tab_sel":   "#3b7ddd",
    "row_move":  "#eaf4fb",
    "row_click": "#eafaf1",
    "row_rel":   "#fefaed",
    "row_key":   "#f4eafb",
    "row_scroll":"#eafbfb",
    "row_sleep": "#f9f9f9",
    "sel":       "#d6e8ff",
}

STYLE = f"""
/* ── Global ── */
QMainWindow, QWidget {{
    background-color: {C['bg']};
    color: {C['text']};
    font-family: 'Segoe UI', 'Microsoft JhengHei', sans-serif;
    font-size: 13px;
}}

/* ── Action bar ── */
#action_bar {{
    background-color: {C['bar_bg']};
}}
#action_bar QLabel {{
    color: {C['bar_text']};
    font-size: 12px;
}}
#action_bar QCheckBox {{
    color: {C['bar_text']};
}}
#action_bar QCheckBox::indicator {{
    border-color: {C['bar_text']};
    background: transparent;
}}
#action_bar QCheckBox::indicator:checked {{
    background: {C['accent']};
    border-color: {C['accent']};
}}

/* ── Action bar buttons ── */
QPushButton#rec_start {{
    background: {C['rec']}; color: white; font-weight: bold;
    border: none; border-radius: 5px;
    min-height: 34px; font-size: 13px; padding: 0 16px;
}}
QPushButton#rec_start:hover   {{ background: {C['rec_hover']}; }}
QPushButton#rec_start:disabled {{ background: {C['rec_dim']}; color: #c0a0a0; }}

QPushButton#rec_stop {{
    background: #95a5a6; color: white; border: none; border-radius: 5px;
    min-height: 34px; font-size: 13px; padding: 0 16px;
}}
QPushButton#rec_stop:hover {{ background: {C['stop_hover']}; }}

QPushButton#play_start {{
    background: {C['play']}; color: white; font-weight: bold;
    border: none; border-radius: 5px;
    min-height: 34px; font-size: 13px; padding: 0 16px;
}}
QPushButton#play_start:hover   {{ background: {C['play_hover']}; }}
QPushButton#play_start:disabled {{ background: {C['play_dim']}; color: #80b890; }}

QPushButton#play_stop {{
    background: #95a5a6; color: white; border: none; border-radius: 5px;
    min-height: 34px; font-size: 13px; padding: 0 16px;
}}
QPushButton#play_stop:hover {{ background: {C['stop_hover']}; }}

/* ── General buttons ── */
QPushButton {{
    background: {C['surface']};
    border: 1px solid {C['border2']};
    border-radius: 5px;
    padding: 4px 12px;
    color: {C['text']};
    min-height: 26px;
}}
QPushButton:hover   {{ background: {C['bg']}; border-color: {C['accent']}; color: {C['accent']}; }}
QPushButton:pressed {{ background: #dce8fa; }}
QPushButton:disabled {{ color: {C['subtle']}; background: {C['bg']}; border-color: {C['border']}; }}

QPushButton#lock_btn:checked {{
    background: #fff3cd; border-color: #f39c12; color: #d68910; font-weight: bold;
}}
QPushButton#danger {{
    color: {C['danger_txt']}; border-color: {C['border']};
}}
QPushButton#danger:hover {{
    background: #fde8e8; border-color: {C['rec']}; color: {C['rec']};
}}

/* ── Tabs ── */
QTabWidget::pane {{
    border: 1px solid {C['border']};
    border-top: none;
    background: {C['surface']};
    border-radius: 0 0 6px 6px;
}}
QTabBar::tab {{
    background: {C['bg']};
    border: 1px solid {C['border']};
    border-bottom: none;
    border-radius: 5px 5px 0 0;
    padding: 6px 18px;
    color: {C['muted']};
    margin-right: 2px;
    font-size: 12px;
}}
QTabBar::tab:selected {{
    background: {C['surface']};
    color: {C['accent']};
    font-weight: bold;
    border-bottom: none;
}}
QTabBar::tab:hover:!selected {{ color: {C['text']}; }}

/* ── Inputs ── */
QComboBox {{
    background: {C['surface']};
    border: 1px solid {C['border2']};
    border-radius: 5px;
    padding: 4px 8px;
    color: {C['text']};
    selection-background-color: {C['sel']};
}}
QComboBox:focus {{ border-color: {C['accent']}; }}
QComboBox::drop-down {{ border: none; width: 20px; }}
QComboBox QAbstractItemView {{
    background: {C['surface']}; border: 1px solid {C['border2']};
    color: {C['text']}; selection-background-color: {C['sel']};
}}
QSpinBox {{
    background: {C['surface']};
    border: 1px solid {C['border2']};
    border-radius: 5px;
    padding: 3px 6px;
    color: {C['text']};
}}
QSpinBox:focus {{ border-color: {C['accent']}; }}
QSpinBox::up-button, QSpinBox::down-button {{ width: 0; border: none; }}

QCheckBox {{ color: {C['text']}; spacing: 6px; }}
QCheckBox::indicator {{
    width: 15px; height: 15px;
    border: 1px solid {C['border2']}; border-radius: 3px;
    background: {C['surface']};
}}
QCheckBox::indicator:checked {{
    background: {C['accent']}; border-color: {C['accent']};
}}

/* ── Table ── */
QTableWidget {{
    background: {C['surface']};
    alternate-background-color: {C['bg']};
    border: 1px solid {C['border']};
    border-radius: 6px;
    gridline-color: {C['border']};
    color: {C['text']};
    selection-background-color: {C['sel']};
    selection-color: {C['text']};
}}
QTableWidget::item {{ padding: 3px 8px; }}
QTableWidget::item:selected {{ background: {C['sel']}; color: {C['text']}; }}
QHeaderView::section {{
    background: {C['bg']};
    border: none;
    border-right: 1px solid {C['border']};
    border-bottom: 1px solid {C['border']};
    padding: 5px 8px;
    color: {C['muted']};
    font-weight: bold;
    font-size: 12px;
}}

/* ── Separator ── */
#sep_h {{ background: {C['border']}; max-height: 1px; }}
#sep_v {{ background: {C['bar_bg']}; max-width: 1px; }}

/* ── Status bar ── */
QStatusBar {{
    background: {C['surface']};
    color: {C['muted']};
    border-top: 1px solid {C['border']};
    font-size: 12px;
}}

/* ── Slider ── */
QSlider::groove:horizontal {{
    height: 4px; background: {C['border2']}; border-radius: 2px;
}}
QSlider::handle:horizontal {{
    background: {C['accent']}; width: 14px; height: 14px;
    margin: -5px 0; border-radius: 7px;
}}
QSlider::sub-page:horizontal {{ background: {C['accent']}; border-radius: 2px; }}

/* ── Splitter ── */
QSplitter::handle:horizontal {{ background: {C['border']}; width: 2px; }}

/* ── Section labels ── */
QLabel#section {{
    color: {C['muted']};
    font-size: 11px;
    font-weight: bold;
    letter-spacing: 1px;
}}
QLabel#hint {{ color: {C['subtle']}; font-size: 11px; }}
QLabel#count {{ color: {C['muted']}; font-size: 12px; font-weight: bold; }}
"""

_ROW_COLORS = {
    "mouse_move":  C["row_move"],
    "click":       C["row_click"],
    "release":     C["row_rel"],
    "scroll":      C["row_scroll"],
    "key":         C["row_key"],
    "sleep":       C["row_sleep"],
    "call_script": "#fff0e6",   # 橘色調，明顯區隔
}

# ---------------------------------------------------------------------------
# Signal bridges
# ---------------------------------------------------------------------------
class RecorderSignals(QObject):
    finished = pyqtSignal(list)

class PlayerSignals(QObject):
    finished = pyqtSignal()
    progress = pyqtSignal(int, object)

class WindowSignals(QObject):
    restored = pyqtSignal()

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
def _sep_h() -> QFrame:
    f = QFrame(); f.setObjectName("sep_h")
    f.setFixedHeight(1)
    f.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Fixed)
    return f


def _btn(text: str, obj: str | None = None, width: int | None = None) -> QPushButton:
    b = QPushButton(text)
    if obj:   b.setObjectName(obj)
    if width: b.setFixedWidth(width)
    return b

def _lbl(text: str, obj: str | None = None) -> QLabel:
    l = QLabel(text)
    if obj: l.setObjectName(obj)
    return l

# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self._wm = WindowManager()
        self._recorder = Recorder()
        self._player = Player()
        self._sm = ScriptManager()
        self._current_events: list[dict] = []
        self._editing = False

        self._setup_ui()
        self._setup_signals()
        self._setup_hotkeys()
        self._refresh_windows()
        self._refresh_scripts()
        self._refresh_window_configs()

    # ------------------------------------------------------------------ #
    #  UI                                                                  #
    # ------------------------------------------------------------------ #
    def _setup_ui(self):
        self.setWindowTitle("DeskPilot")
        self.setMinimumSize(820, 520)
        self.resize(1000, 620)
        self.setWindowFlag(Qt.WindowType.WindowStaysOnTopHint, True)
        self.setStyleSheet(STYLE)

        root = QWidget()
        self.setCentralWidget(root)
        rl = QVBoxLayout(root)
        rl.setContentsMargins(0, 0, 0, 0)
        rl.setSpacing(0)

        rl.addWidget(self._build_action_bar())

        body = QWidget()
        body.setStyleSheet(f"background:{C['bg']};")
        bl = QHBoxLayout(body)
        bl.setContentsMargins(10, 10, 10, 10)
        bl.setSpacing(10)

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setChildrenCollapsible(False)
        splitter.addWidget(self._build_left_panel())
        splitter.addWidget(self._build_event_panel())
        splitter.setSizes([300, 680])
        bl.addWidget(splitter)
        rl.addWidget(body, 1)

        # Status bar
        self._status = QStatusBar()
        self.setStatusBar(self._status)
        self._status.showMessage("就緒   F6 播放   F7 停止   F8 開始錄製   F9 停止錄製")

    # ── Action bar ─────────────────────────────────────────────────────
    def _build_action_bar(self) -> QWidget:
        bar = QWidget()
        bar.setObjectName("action_bar")
        bar.setFixedHeight(58)
        bl = QHBoxLayout(bar)
        bl.setContentsMargins(14, 10, 14, 10)
        bl.setSpacing(8)

        self._btn_rec_start = _btn("⏺  開始錄製  F8", "rec_start")
        self._btn_rec_stop  = _btn("⏹  停止錄製  F9", "rec_stop")
        self._btn_rec_stop.setEnabled(False)
        self._btn_play = _btn("▶  播放  F6", "play_start")
        self._btn_stop = _btn("⏹  停止  F7", "play_stop")
        self._btn_stop.setEnabled(False)

        bl.addWidget(self._btn_rec_start)
        bl.addWidget(self._btn_rec_stop)
        bl.addSpacing(12)
        bl.addWidget(self._btn_play)
        bl.addWidget(self._btn_stop)
        bl.addStretch()
        return bar

    # ── Left panel ─────────────────────────────────────────────────────
    def _build_left_panel(self) -> QTabWidget:
        tabs = QTabWidget()
        tabs.setMaximumWidth(320)
        tabs.addTab(self._build_window_tab(), "  視窗控制  ")
        tabs.addTab(self._build_script_tab(), "  腳本管理  ")
        return tabs

    def _build_window_tab(self) -> QWidget:
        w = QWidget()
        w.setStyleSheet(f"background:{C['surface']};")
        layout = QVBoxLayout(w)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        # 目標視窗
        layout.addWidget(_lbl("目標視窗", "section"))
        row = QHBoxLayout()
        self._win_combo = QComboBox()
        self._btn_refresh = _btn("重新整理")
        self._btn_refresh.setToolTip("重新整理視窗清單")
        row.addWidget(self._win_combo, 1)
        row.addWidget(self._btn_refresh)
        layout.addLayout(row)

        layout.addWidget(_sep_h())

        # 位置與尺寸
        layout.addWidget(_lbl("位置與尺寸", "section"))
        grid = QHBoxLayout()
        grid.setSpacing(8)
        for text, attr in [("X", "_spin_x"), ("Y", "_spin_y"),
                            ("W", "_spin_w"), ("H", "_spin_h")]:
            col = QVBoxLayout()
            col.setSpacing(3)
            lbl = QLabel(text)
            lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
            lbl.setStyleSheet(f"color:{C['muted']}; font-size:11px; font-weight:bold;")
            spin = QSpinBox()
            spin.setRange(-9999, 9999)
            setattr(self, attr, spin)
            col.addWidget(lbl)
            col.addWidget(spin)
            grid.addLayout(col)
        layout.addLayout(grid)

        row2 = QHBoxLayout()
        row2.setSpacing(6)
        self._chk_topmost = QCheckBox("目標視窗置頂")
        self._btn_apply = _btn("套用")
        self._btn_lock  = _btn("🔒 鎖定", "lock_btn")
        self._btn_lock.setCheckable(True)
        row2.addWidget(self._chk_topmost)
        row2.addStretch()
        row2.addWidget(self._btn_apply)
        row2.addWidget(self._btn_lock)
        layout.addLayout(row2)

        layout.addWidget(_sep_h())

        # 視窗設定
        layout.addWidget(_lbl("視窗設定", "section"))
        row3 = QHBoxLayout()
        row3.setSpacing(6)
        self._wincfg_combo  = QComboBox()
        self._btn_wincfg_save   = _btn("儲存")
        self._btn_wincfg_load   = _btn("載入")
        self._btn_wincfg_delete = _btn("✕", "danger", width=30)
        row3.addWidget(self._wincfg_combo, 1)
        row3.addWidget(self._btn_wincfg_save)
        row3.addWidget(self._btn_wincfg_load)
        row3.addWidget(self._btn_wincfg_delete)
        layout.addLayout(row3)

        layout.addStretch()
        return w

    def _build_script_tab(self) -> QWidget:
        w = QWidget()
        w.setStyleSheet(f"background:{C['surface']};")
        layout = QVBoxLayout(w)
        layout.setContentsMargins(14, 14, 14, 14)
        layout.setSpacing(12)

        layout.addWidget(_lbl("腳本", "section"))
        self._script_combo = QComboBox()
        layout.addWidget(self._script_combo)

        row = QHBoxLayout()
        row.setSpacing(6)
        self._btn_save      = _btn("儲存")
        self._btn_load      = _btn("載入")
        self._btn_load_file = _btn("開啟檔案")
        self._btn_delete    = _btn("刪除", "danger")
        for b in [self._btn_save, self._btn_load, self._btn_load_file, self._btn_delete]:
            row.addWidget(b)
        layout.addLayout(row)

        layout.addWidget(_sep_h())

        layout.addWidget(_lbl("播放設定", "section"))
        repeat_row = QHBoxLayout()
        repeat_row.setSpacing(8)
        repeat_row.addWidget(_lbl("次數"))
        self._spin_repeat = QSpinBox()
        self._spin_repeat.setRange(1, 9999)
        self._spin_repeat.setValue(1)
        self._spin_repeat.setFixedWidth(72)
        self._chk_loop = QCheckBox("無限循環")
        repeat_row.addWidget(self._spin_repeat)
        repeat_row.addWidget(self._chk_loop)
        repeat_row.addStretch()
        layout.addLayout(repeat_row)

        layout.addStretch()
        return w

    # ── Event panel ────────────────────────────────────────────────────
    def _build_event_panel(self) -> QWidget:
        w = QWidget()
        layout = QVBoxLayout(w)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(6)

        # Toolbar
        tb = QHBoxLayout()
        tb.setSpacing(6)
        self._lbl_event_count = _lbl("操作序列（0 個事件）", "count")
        self._btn_ev_up     = _btn("上移")
        self._btn_ev_down   = _btn("下移")
        self._btn_ev_call   = _btn("插入呼叫腳本")
        self._btn_ev_delete = _btn("刪除列", "danger")
        self._btn_ev_clear  = _btn("清除全部", "danger")
        for b in [self._btn_ev_up, self._btn_ev_down, self._btn_ev_call,
                  self._btn_ev_delete, self._btn_ev_clear]:
            b.setFixedHeight(26)
        tb.addWidget(self._lbl_event_count)
        tb.addStretch()
        tb.addWidget(self._btn_ev_up)
        tb.addWidget(self._btn_ev_down)
        tb.addWidget(self._btn_ev_call)
        tb.addWidget(self._btn_ev_delete)
        tb.addWidget(self._btn_ev_clear)
        layout.addLayout(tb)

        # Table
        self._event_table = QTableWidget()
        self._event_table.setColumnCount(4)
        self._event_table.setHorizontalHeaderLabels(["#", "類型", "參數1", "參數2"])
        self._event_table.setAlternatingRowColors(False)
        self._event_table.setShowGrid(True)
        self._event_table.verticalHeader().setVisible(False)
        self._event_table.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectRows)
        self._event_table.setEditTriggers(
            QTableWidget.EditTrigger.DoubleClicked |
            QTableWidget.EditTrigger.SelectedClicked
        )
        hdr = self._event_table.horizontalHeader()
        hdr.setSectionResizeMode(0, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(1, QHeaderView.ResizeMode.Fixed)
        hdr.setSectionResizeMode(2, QHeaderView.ResizeMode.Stretch)
        hdr.setSectionResizeMode(3, QHeaderView.ResizeMode.Stretch)
        self._event_table.setColumnWidth(0, 38)
        self._event_table.setColumnWidth(1, 92)
        self._event_table.setStyleSheet(
            f"QTableWidget {{ border-radius: 6px; }}"
        )
        layout.addWidget(self._event_table, 1)

        hint = _lbl(
            "雙擊儲存格可編輯  ·  sleep 參數1 = 毫秒數  ·  "
            "mouse 參數1 = X，參數2 = Y  ·  key 參數1 = 按鍵字元",
            "hint"
        )
        layout.addWidget(hint)
        return w

    # ------------------------------------------------------------------ #
    #  Signals                                                             #
    # ------------------------------------------------------------------ #
    def _setup_signals(self):
        self._btn_refresh.clicked.connect(self._refresh_windows)
        self._win_combo.currentIndexChanged.connect(self._on_window_selected)
        self._btn_apply.clicked.connect(self._apply_window_pos)
        self._btn_lock.toggled.connect(self._toggle_lock)
        self._chk_topmost.toggled.connect(lambda v: self._wm.set_always_on_top(v))
        for s in [self._spin_x, self._spin_y, self._spin_w, self._spin_h]:
            s.valueChanged.connect(self._on_spin_changed)

        self._btn_wincfg_save.clicked.connect(self._save_window_config)
        self._btn_wincfg_load.clicked.connect(self._load_window_config)
        self._btn_wincfg_delete.clicked.connect(self._delete_window_config)

        self._rec_signals = RecorderSignals()
        self._rec_signals.finished.connect(self._on_record_finished)
        self._btn_rec_start.clicked.connect(self._start_recording)
        self._btn_rec_stop.clicked.connect(self._stop_recording)

        self._play_signals = PlayerSignals()
        self._play_signals.finished.connect(self._on_play_finished)
        self._play_signals.progress.connect(self._on_play_progress)
        self._player.set_on_finished(lambda: self._play_signals.finished.emit())
        self._player.set_on_progress(lambda c, t: self._play_signals.progress.emit(c, t))
        self._player.set_script_loader(self._sm.load_script)
        self._btn_play.clicked.connect(self._start_play)
        self._btn_stop.clicked.connect(self._stop_play)

        self._win_signals = WindowSignals()
        self._win_signals.restored.connect(
            lambda: self._status.showMessage("視窗位置已自動還原"))
        self._wm.set_on_restored_callback(lambda: self._win_signals.restored.emit())

        self._btn_save.clicked.connect(self._save_script)
        self._btn_load.clicked.connect(self._load_script)
        self._btn_load_file.clicked.connect(self._load_script_file)
        self._btn_delete.clicked.connect(self._delete_script)
        self._chk_loop.toggled.connect(lambda v: self._spin_repeat.setEnabled(not v))

        self._event_table.cellChanged.connect(self._on_cell_changed)
        self._btn_ev_up.clicked.connect(self._move_event_up)
        self._btn_ev_down.clicked.connect(self._move_event_down)
        self._btn_ev_call.clicked.connect(self._insert_call_script)
        self._btn_ev_delete.clicked.connect(self._delete_event_row)
        self._btn_ev_clear.clicked.connect(self._clear_events)

    def _setup_hotkeys(self):
        QShortcut(QKeySequence("F6"), self).activated.connect(self._start_play)
        QShortcut(QKeySequence("F7"), self).activated.connect(self._stop_play)
        QShortcut(QKeySequence("F8"), self).activated.connect(self._start_recording)
        QShortcut(QKeySequence("F9"), self).activated.connect(self._stop_recording)

    # ------------------------------------------------------------------ #
    #  Window control                                                      #
    # ------------------------------------------------------------------ #
    def _refresh_windows(self):
        self._win_combo.blockSignals(True)
        self._win_combo.clear()
        for w in self._wm.get_window_list():
            self._win_combo.addItem(w["title"], userData=w["hwnd"])
        self._win_combo.blockSignals(False)
        if self._win_combo.count() > 0:
            self._on_window_selected(0)

    def _on_window_selected(self, index: int):
        if index < 0:
            return
        hwnd = self._win_combo.itemData(index)
        if hwnd:
            self._wm.select_window(hwnd)
            rect = self._wm.get_current_rect()
            if rect:
                self._update_spin_from_rect(rect)

    def _update_spin_from_rect(self, rect: dict):
        for spin, key in [(self._spin_x, "x"), (self._spin_y, "y"),
                          (self._spin_w, "w"), (self._spin_h, "h")]:
            spin.blockSignals(True)
            spin.setValue(rect[key])
            spin.blockSignals(False)

    def _on_spin_changed(self):
        self._wm.set_target(
            self._spin_x.value(), self._spin_y.value(),
            self._spin_w.value(), self._spin_h.value(),
        )

    def _apply_window_pos(self):
        self._wm.set_target(
            self._spin_x.value(), self._spin_y.value(),
            self._spin_w.value(), self._spin_h.value(),
        )
        ok = self._wm.apply_position()
        self._status.showMessage("位置已套用" if ok else "套用失敗：未選取有效視窗")

    def _toggle_lock(self, checked: bool):
        if checked:
            self._wm.set_target(
                self._spin_x.value(), self._spin_y.value(),
                self._spin_w.value(), self._spin_h.value(),
            )
            self._wm.lock()
            self._btn_lock.setText("🔓 解鎖")
            self._status.showMessage("視窗已鎖定")
        else:
            self._wm.unlock()
            self._btn_lock.setText("🔒 鎖定")
            self._status.showMessage("視窗已解鎖")

    # ------------------------------------------------------------------ #
    #  Window config                                                       #
    # ------------------------------------------------------------------ #
    def _refresh_window_configs(self):
        self._wincfg_combo.clear()
        for name in self._sm.list_window_configs():
            self._wincfg_combo.addItem(name)

    def _save_window_config(self):
        name, ok = QInputDialog.getText(self, "儲存視窗設定", "設定名稱：")
        if ok and name.strip():
            name = name.strip()
            config = self._wm.get_config()
            config["window_title"] = self._win_combo.currentText()
            self._sm.save_window_config(name, config)
            self._refresh_window_configs()
            idx = self._wincfg_combo.findText(name)
            if idx >= 0:
                self._wincfg_combo.setCurrentIndex(idx)
            self._status.showMessage(f"視窗設定「{name}」已儲存")

    def _load_window_config(self):
        name = self._wincfg_combo.currentText()
        if not name:
            return
        config = self._sm.load_window_config(name)
        if not config:
            self._status.showMessage("找不到設定檔")
            return
        self._wm.load_config(config)
        target = config.get("target", {})
        if target:
            self._update_spin_from_rect(
                {"x": target["x"], "y": target["y"],
                 "w": target["w"], "h": target["h"]}
            )
        self._chk_topmost.setChecked(config.get("always_on_top", False))
        self._status.showMessage(f"視窗設定「{name}」已載入")

    def _delete_window_config(self):
        name = self._wincfg_combo.currentText()
        if not name:
            return
        reply = QMessageBox.question(
            self, "確認刪除", f"確定要刪除視窗設定「{name}」？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self._sm.delete_window_config(name)
            self._refresh_window_configs()
            self._status.showMessage(f"視窗設定「{name}」已刪除")

    # ------------------------------------------------------------------ #
    #  Recording                                                           #
    # ------------------------------------------------------------------ #
    def _start_recording(self):
        if self._recorder.is_recording():
            return
        self._recorder.start()
        self._btn_rec_start.setEnabled(False)
        self._btn_rec_stop.setEnabled(True)
        self._status.showMessage("● 錄製中...  按 F9 或點擊「停止錄製」結束")

    def _stop_recording(self):
        if not self._recorder.is_recording():
            return
        events = self._recorder.stop()
        self._btn_rec_start.setEnabled(True)
        self._btn_rec_stop.setEnabled(False)
        self._rec_signals.finished.emit(events)

    def _on_record_finished(self, events: list[dict]):
        self._current_events = events
        self._populate_event_table(events)
        self._status.showMessage(f"錄製完成，共 {len(events)} 個事件")

    # ------------------------------------------------------------------ #
    #  Playback                                                            #
    # ------------------------------------------------------------------ #
    def _start_play(self):
        if not self._current_events:
            self._status.showMessage("無可播放的腳本，請先錄製或載入")
            return
        if self._player.is_playing():
            return
        self._player.play(
            self._current_events,
            repeat=self._spin_repeat.value(),
            loop=self._chk_loop.isChecked(),
        )
        self._btn_play.setEnabled(False)
        self._btn_stop.setEnabled(True)
        self._status.showMessage("▶ 播放中...")

    def _stop_play(self):
        self._player.stop()

    def _on_play_finished(self):
        self._btn_play.setEnabled(True)
        self._btn_stop.setEnabled(False)
        self._status.showMessage("播放完成")

    def _on_play_progress(self, current: int, total):
        msg = (f"▶ 播放中... 第 {current} 次（無限循環）"
               if total is None else f"▶ 播放中... 第 {current} / {total} 次")
        self._status.showMessage(msg)

    # ------------------------------------------------------------------ #
    #  Script management                                                   #
    # ------------------------------------------------------------------ #
    def _refresh_scripts(self):
        self._script_combo.clear()
        for name in self._sm.list_scripts():
            self._script_combo.addItem(name)

    def _save_script(self):
        if not self._current_events:
            self._status.showMessage("無腳本可儲存")
            return
        name, ok = QInputDialog.getText(self, "儲存腳本", "腳本名稱：")
        if ok and name.strip():
            name = name.strip()
            self._sm.save_script(name, self._current_events)
            self._refresh_scripts()
            idx = self._script_combo.findText(name)
            if idx >= 0:
                self._script_combo.setCurrentIndex(idx)
            self._status.showMessage(f"腳本「{name}」已儲存")

    def _load_script(self):
        name = self._script_combo.currentText()
        if not name:
            return
        try:
            events = self._sm.load_script(name)
            self._current_events = events
            self._populate_event_table(events)
            self._status.showMessage(f"腳本「{name}」已載入，共 {len(events)} 個事件")
        except Exception as e:
            QMessageBox.critical(self, "載入失敗", str(e))

    def _load_script_file(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "開啟腳本", str(self._sm._dir), "JSON Files (*.json)"
        )
        if path:
            try:
                events = self._sm.load_script_from_path(path)
                self._current_events = events
                self._populate_event_table(events)
                self._status.showMessage(f"已載入：{Path(path).name}，共 {len(events)} 個事件")
            except Exception as e:
                QMessageBox.critical(self, "載入失敗", str(e))

    def _delete_script(self):
        name = self._script_combo.currentText()
        if not name:
            return
        reply = QMessageBox.question(
            self, "確認刪除", f"確定要刪除腳本「{name}」？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self._sm.delete_script(name)
            self._refresh_scripts()
            self._status.showMessage(f"腳本「{name}」已刪除")

    # ------------------------------------------------------------------ #
    #  Event table                                                         #
    # ------------------------------------------------------------------ #
    def _populate_event_table(self, events: list[dict]):
        self._editing = True
        self._event_table.setRowCount(0)
        for i, ev in enumerate(events):
            row = self._event_table.rowCount()
            self._event_table.insertRow(row)
            etype = ev.get("type", "")
            p1, p2 = self._event_params(ev)

            idx_item = QTableWidgetItem(str(i + 1))
            idx_item.setTextAlignment(Qt.AlignmentFlag.AlignCenter)
            idx_item.setFlags(idx_item.flags() & ~Qt.ItemFlag.ItemIsEditable)
            idx_item.setForeground(QColor(C["muted"]))
            self._event_table.setItem(row, 0, idx_item)

            for col, val in enumerate([etype, p1, p2], start=1):
                self._event_table.setItem(row, col, QTableWidgetItem(val))

            bg = QColor(_ROW_COLORS.get(etype, C["surface"]))
            for col in range(4):
                item = self._event_table.item(row, col)
                if item:
                    item.setBackground(bg)

        self._lbl_event_count.setText(f"操作序列（{len(events)} 個事件）")
        self._editing = False

    @staticmethod
    def _event_params(ev: dict) -> tuple[str, str]:
        etype = ev.get("type", "")
        if etype in ("mouse_move", "click", "release", "scroll"):
            p1 = str(ev.get("x", ""))
            p2 = str(ev.get("y", ""))
            if "button" in ev:
                p2 += f"  btn={ev['button']}"
        elif etype == "key":
            p1 = ev.get("key", "")
            p2 = ""
        elif etype == "sleep":
            p1 = str(ev.get("ms", 0))
            p2 = ""
        elif etype == "call_script":
            p1 = ev.get("name", "")
            p2 = ""
        else:
            p1 = p2 = ""
        return p1, p2

    def _on_cell_changed(self, row: int, col: int):
        if self._editing or col == 0 or row >= len(self._current_events):
            return
        item = self._event_table.item(row, col)
        if not item:
            return
        value = item.text().strip()
        ev = self._current_events[row]
        if col == 1:
            ev["type"] = value
        elif col == 2:
            self._apply_p1(ev, value)
        elif col == 3:
            self._apply_p2(ev, value)

    @staticmethod
    def _apply_p1(ev: dict, value: str):
        etype = ev.get("type", "")
        if etype == "sleep":
            try:
                ev["ms"] = int(value.replace("ms", "").strip())
            except ValueError:
                pass
        elif etype == "key":
            ev["key"] = value
        elif etype in ("mouse_move", "click", "release", "scroll"):
            try:
                ev["x"] = int(value)
            except ValueError:
                pass

    @staticmethod
    def _apply_p2(ev: dict, value: str):
        if ev.get("type") in ("mouse_move", "click", "release", "scroll"):
            for part in value.split():
                if part.startswith("btn="):
                    ev["button"] = part[4:]
                else:
                    try:
                        ev["y"] = int(part)
                    except ValueError:
                        pass

    def _selected_row(self) -> int:
        items = self._event_table.selectedItems()
        return self._event_table.row(items[0]) if items else -1

    def _delete_event_row(self):
        row = self._selected_row()
        if row < 0 or row >= len(self._current_events):
            return
        self._current_events.pop(row)
        self._populate_event_table(self._current_events)
        new_row = min(row, len(self._current_events) - 1)
        if new_row >= 0:
            self._event_table.selectRow(new_row)

    def _move_event_up(self):
        row = self._selected_row()
        if row <= 0 or row >= len(self._current_events):
            return
        evs = self._current_events
        evs[row - 1], evs[row] = evs[row], evs[row - 1]
        self._populate_event_table(evs)
        self._event_table.selectRow(row - 1)

    def _move_event_down(self):
        row = self._selected_row()
        if row < 0 or row >= len(self._current_events) - 1:
            return
        evs = self._current_events
        evs[row], evs[row + 1] = evs[row + 1], evs[row]
        self._populate_event_table(evs)
        self._event_table.selectRow(row + 1)

    def _insert_call_script(self):
        scripts = self._sm.list_scripts()
        if not scripts:
            QMessageBox.information(self, "無可用腳本", "請先儲存至少一個腳本再使用呼叫功能。")
            return
        name, ok = QInputDialog.getItem(
            self, "插入呼叫腳本", "選擇要呼叫的腳本：", scripts, 0, False
        )
        if not ok or not name:
            return
        event = {"type": "call_script", "name": name}
        row = self._selected_row()
        if row >= 0:
            self._current_events.insert(row + 1, event)
        else:
            self._current_events.append(event)
        self._populate_event_table(self._current_events)

    def _clear_events(self):
        if not self._current_events:
            return
        reply = QMessageBox.question(
            self, "確認清除", "確定要清除所有事件？",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if reply == QMessageBox.StandardButton.Yes:
            self._current_events = []
            self._populate_event_table([])

    def closeEvent(self, event):
        self._player.stop()
        self._wm.unlock()
        if self._recorder.is_recording():
            self._recorder.stop()
        super().closeEvent(event)
