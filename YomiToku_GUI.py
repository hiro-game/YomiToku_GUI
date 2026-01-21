# ============================================================
# 1. import 群
# ============================================================
import sys
import subprocess
import configparser
import importlib
from pathlib import Path
INI_FILE = "YomiToku_GUI.ini"

from PySide6.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QLabel,
    QFileDialog, QTextEdit, QComboBox, QCheckBox,
    QLineEdit, QHBoxLayout, QListWidget, QGridLayout
)
from PySide6.QtGui import QIntValidator, QFontMetrics, QFont
from PySide6.QtCore import Qt, QThread, QObject, Signal
from PySide6.QtWidgets import QProxyStyle, QStyle
# switch_widget.py
from PySide6.QtWidgets import QWidget
from PySide6.QtCore import Qt, QRect, QSize
from PySide6.QtGui import QColor, QPainter, QBrush, QPen
class SwitchWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self._checked = False
        self.setFixedSize(42, 20)

    def sizeHint(self):
        return QSize(42, 20)

    def isChecked(self):
        return self._checked

    def setChecked(self, value: bool):
        self._checked = bool(value)
        self.update()

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._checked = not self._checked
            self.update()

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 背景色
        bg_color = QColor(160, 160, 160) if not self._checked else QColor(210, 210, 210)
        painter.setBrush(QBrush(bg_color))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(self.rect(), 4, 4)

        # ハンドル
        handle_size = 16
        y = (self.height() - handle_size) // 2
        x = self.width() - handle_size - 2 if self._checked else 2

        handle_rect = QRect(x, y, handle_size, handle_size)

        painter.setBrush(QBrush(QColor(240, 240, 240)))
        painter.setPen(QPen(QColor(150, 150, 150), 1))
        painter.drawRoundedRect(handle_rect, 4, 4)

# ============================================================
# 2. Worker クラス（バックエンド処理）
# ============================================================
class YomiTokuWorker(QObject):
    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished = Signal()

    def __init__(
            self,
            exe_path,
            input_paths,
            outdir_path,
            fmt,
            dpi,
            reading_order,
            pages,
            figure,
            table,
            lite,
            vis
    ):
            super().__init__()
            self.exe_path = exe_path
            self.input_paths = input_paths
            self.outdir_path = outdir_path
            self.fmt = fmt
            self.device = device
            self.dpi = dpi
            self.reading_order = reading_order
            self.pages = pages
            self.figure = figure
            self.table = table
            self.lite = lite
            self.vis = vis

    def run(self):
            total = len(self.input_paths)
            self.log_signal.emit(f"----- {total} 件の処理を開始 -----")

            for idx, input_path in enumerate(self.input_paths, start=1):
                    progress = int((idx - 1) / total * 100)
                    self.progress_signal.emit(progress)

                    self.log_signal.emit(f"[{idx}/{total}] 処理中: {input_path}")

                    cmd = [
                            str(self.exe_path),
                            str(input_path),
                            "-o", str(self.outdir_path),
                            "-f", self.fmt,
                            "-d", self.device
                    ]

                    if self.dpi:
                            cmd += ["--dpi", self.dpi]
                    if self.reading_order:
                            cmd += ["--reading_order", self.reading_order]
                    if self.pages:
                            cmd += ["--pages", self.pages]
                    if self.figure:
                            cmd.append("--figure")
                    if self.table:
                            cmd.append("--table")
                    if self.lite:
                            cmd.append("--lite")
                    if self.vis:
                            cmd.append("--vis")

                    process = subprocess.Popen(
                            cmd,
                            stdout=subprocess.PIPE,
                            stderr=subprocess.STDOUT,
                            text=True
                    )

                    for line in process.stdout:
                            self.log_signal.emit(line.rstrip())

                    process.wait()
                    self.log_signal.emit(
                            f"完了: {input_path} (終了コード: {process.returncode})"
                    )

            self.progress_signal.emit(100)
            self.log_signal.emit("----- 全ての処理が完了しました -----")
            self.finished.emit()

# ============================================================
# 3. GUI 本体
# ============================================================
class YomiTokuGUI(QWidget):
    SUPPORTED_EXT = [".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"]

# --------------------------------------------------------
# 3-1. 初期化
# --------------------------------------------------------
    def __init__(self):
        super().__init__()

        font = QFont()
        font.setPointSize(12)
        self.setFont(font)

        self._init_basic_state()
        self._build_ui()
        self._load_initial_config()

        # ★ 設定を UI に反映（これが抜けていた）
        self.load_all_settings()

        # ★ YomiToku のパスを読み込む
        self.load_yomitoku_path()

        self.adjustSize()
        self.setFixedSize(self.size())

        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log(f"=== App started at {timestamp} ===")

    # --------------------------------------------------------
    # ★ YomiToku のパスを読み込む（初回は自動取得）
    #    - 設定ファイルに有効なパスがあればそれを最優先
    #    - 無ければ自動検出（site → where/which → ~/.local/bin）
    # --------------------------------------------------------
    def load_yomitoku_path(self):
        cfg = self.config

        if "Settings" not in cfg:
            cfg["Settings"] = {}

        settings = cfg["Settings"]

        # 1. 設定ファイルにパスがある場合はそれを最優先（ユーザーの意思を尊重）
        # 1. 設定ファイルにパスがある場合はそれを最優先（ユーザーの意思を尊重）
        if "yomitoku_path" in settings:
            raw = settings["yomitoku_path"].strip()

            # 設定ファイルに値が入っている場合
            if raw:
                path = Path(raw)

                # ★ 有効なパスなら即採用（自動検出は絶対に行わない）
                if path.exists():
                    self.yomitoku_path = path
                    return

                # パスが存在しない場合のみ自動検出へ進む
                self.log(f"設定ファイルの YomiToku パスが存在しません: {raw}")

            else:
                self.log("設定ファイルに yomitoku_path が空で保存されています。")

        else:
            self.log("設定ファイルに yomitoku_path がありません。")

        # 2. 自動検出（site.getsitepackages / getusersitepackages → Scripts/bin → where/which）
        auto_path = self.find_yomitoku_exe()
        if auto_path and auto_path.exists():
            self.yomitoku_path = auto_path

            # 設定にまだ値が無い場合のみ保存（既存のユーザー設定は絶対に上書きしない）
            if not settings.get("yomitoku_path", "").strip():
                settings["yomitoku_path"] = str(auto_path)
                self.save_config()
                self.log(f"YomiToku パスを自動検出し、設定に保存しました: {auto_path}")
            else:
                self.log(f"YomiToku パスを自動検出しました（設定は既存値を維持）: {auto_path}")
            return

        # 3. 取得できなかった場合
        self.yomitoku_path = None
        self.log("YomiToku のパスを自動検出できませんでした。設定画面から手動で指定してください。")

    # --------------------------------------------------------
    # ★ YomiToku 実行ファイルの自動検出
    #    - OS 判定なしで site.getsitepackages / getusersitepackages を利用
    #    - Scripts / bin を総当たり
    #    - where / which をフォールバックとして使用
    #    - 複数見つかった場合は「より新しい Python バージョン」を優先
    # --------------------------------------------------------
    def find_yomitoku_exe(self):
        import sys
        import shutil
        import site
        import os
        import re
        from pathlib import Path

        candidates = []

        def add_candidate(p: Path):
            if p and p.exists():
                p = p.resolve()
                if p not in candidates:
                    candidates.append(p)

        # 1. site.getsitepackages / getusersitepackages から Scripts / bin を総当たり
        bases = []

        try:
            bases.extend(site.getsitepackages())
        except Exception:
            pass

        try:
            user_site = site.getusersitepackages()
            if user_site:
                bases.append(user_site)
        except Exception:
            pass

        for base in bases:
            base_path = Path(base).resolve()
            parent = base_path.parent

            for scripts_name in ("Scripts", "bin"):
                folder = parent / scripts_name
                if not folder.exists():
                    continue

                if sys.platform.startswith("win"):
                    exe = folder / "yomitoku.exe"
                    if exe.exists():
                        add_candidate(exe)
                else:
                    script = folder / "yomitoku"
                    if script.exists():
                        add_candidate(script)

        # 2. where / which をフォールバックとして使用
        exe_name = "yomitoku.exe" if sys.platform.startswith("win") else "yomitoku"
        exe = shutil.which(exe_name)
        if exe:
            add_candidate(Path(exe))

        # 3. Unix 系の ~/.local/bin/yomitoku もチェック
        if not sys.platform.startswith("win"):
            local_bin = Path.home() / ".local" / "bin" / "yomitoku"
            if local_bin.exists():
                add_candidate(local_bin)

        if not candidates:
            return None

        # 4. 複数見つかった場合は「Python バージョンが新しいもの」を優先
        def version_key(p: Path):
            s = str(p)

            # 例: Python313, python311 など
            m = re.search(r"[Pp]ython(?:3)?(\d)(\d)", s)
            if m:
                major = int(m.group(1))
                minor = int(m.group(2))
                return (major, minor)

            # 例: python3.11, Python3.10 など
            m = re.search(r"[Pp]ython(\d)\.(\d+)", s)
            if m:
                major = int(m.group(1))
                minor = int(m.group(2))
                return (major, minor)

            # バージョン情報が取れない場合は最低優先
            return (0, 0)

        candidates.sort(key=version_key, reverse=True)
        best = candidates[0]
        return best

    # --------------------------------------------------------
    # ★ ログ表示
    # --------------------------------------------------------
    def log(self, text):
        self.log_view.append(text)

    # --------------------------------------------------------
    # 3-2. 状態・設定関連
    # --------------------------------------------------------
    def _init_basic_state(self):
        self.setWindowTitle("YomiToku_GUI")
        self.resize(900, 650)
        self.setAcceptDrops(True)

        self.config_path = Path("YomiToku_GUI.ini")
        self.config = configparser.ConfigParser()

        self.input_paths = []
        self.output_dir = None

    def _load_initial_config(self):
        self.load_config()

        if "Settings" in self.config and "output_dir" in self.config["Settings"]:
            self.output_dir = Path(self.config["Settings"]["output_dir"])

        # Settings セクション（保存フラグ）
        self.save_settings_flag = self.config.get("Settings", "save_settings", fallback="1") == "1"
        self.save_log_flag = self.config.get("Settings", "save_log", fallback="0") == "1"

    def load_config(self):
        if self.config_path.exists():
            self.config.read(self.config_path, encoding="utf-8")
        else:
            self.config["Settings"] = {
                 "save_settings": "0",
                 "save_log": "0"
            }
            self.save_config()

    def save_config(self):
        with open(self.config_path, "w", encoding="utf-8") as f:
            self.config.write(f)

    def save_all_settings(self):
        cfg = self.config

        if "Settings" not in cfg:
            cfg["Settings"] = {}

        s = cfg["Settings"]
        s["save_settings"] = "1" if self.save_settings_flag else "0"
        s["save_log"] = "1" if self.save_log_flag else "0"

        # 中部の設定内容
        s["format"] = str(self.format_box.currentIndex())
        s["reading_order"] = str(self.direction_box.currentIndex())
        s["dpi"] = self.dpi_box.currentText()
        s["pages"] = self.pages_input.text()
        s["figure_width"] = self.figure_width_input.text()
        s["figure_dir"] = self.figure_dir_input.text()

        # チェックボックス類（UI の変数名に完全一致）
        s["figure"] = "1" if self.figure_check.isChecked() else "0"
        s["table"] = "1" if self.table_check.isChecked() else "0"
        s["lite"] = "1" if self.lite_check.isChecked() else "0"
        s["vis"] = "1" if self.vis_check.isChecked() else "0"
        s["figure_letter"] = "1" if self.figure_letter_check.isChecked() else "0"

        # ignore_line_break / combine / ignore_meta は UI に存在しないため削除
        # 必要なら UI に追加してから保存処理を復活させる

        self.save_config()

    def load_all_settings(self):
        cfg = self.config
        if "Settings" not in cfg:
            return

        s = cfg["Settings"]

        # 中部の設定内容
        if "format" in s:
            self.format_box.setCurrentIndex(int(s["format"]))

        if "reading_order" in s:
            self.direction_box.setCurrentIndex(int(s["reading_order"]))

        if "dpi" in s:
            self.dpi_box.setCurrentText(s["dpi"])

        if "pages" in s:
            self.pages_input.setText(s["pages"])

        if "figure_width" in s:
            self.figure_width_input.setText(s["figure_width"])

        if "figure_dir" in s:
            self.figure_dir_input.setText(s["figure_dir"])

        # チェックボックス類
        if "figure" in s:
            self.figure_check.setChecked(s["figure"] == "1")

        if "table" in s:
            self.table_check.setChecked(s["table"] == "1")

        if "lite" in s:
            self.lite_check.setChecked(s["lite"] == "1")

        if "vis" in s:
            self.vis_check.setChecked(s["vis"] == "1")

        if "figure_letter" in s:
            self.figure_letter_check.setChecked(s["figure_letter"] == "1")

    def closeEvent(self, event):
        # 設定保存（ini の save_settings=1 のときだけ）
        if self.save_settings_flag:
            self.save_all_settings()

        # ログ保存（ini の save_log=1 のときだけ）
        if self.save_log_flag:
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_name = f"YomiToku_{timestamp}.log"
            with open(log_name, "w", encoding="utf-8") as f:
                f.write(self.log_view.toPlainText())

        event.accept()

    # --------------------------------------------------------
    # 3-3. UI 構築（SDI）
    # --------------------------------------------------------
    def _build_ui(self):
        """
        UI 全体を構築する。
        上部 → 中部 → 下部 の順に積み上げる。
        """
        layout = QVBoxLayout()

        # 上部（ファイル選択・リスト）
        self._build_top_section(layout)

        # 中部（各種設定）
        self._build_middle_section(layout)

        # 下部（実行ボタン・ログ）
        self._build_bottom_section(layout)

        self.setLayout(layout)

    # --------------------------------------------------------
    # 3-3-1. 上部 UI（ファイル選択・フォルダ選択・出力先選択・リスト）
    # --------------------------------------------------------
    def _build_top_section(self, parent_layout):
        """
        上部エリア:
        - ファイル選択ボタン
        - フォルダ選択ボタン
        - 出力先選択ボタン
        - 選択されたファイル一覧リスト
        """
        input_area = QHBoxLayout()
        input_area.setSpacing(8)

        left_buttons = QVBoxLayout()
        left_buttons.setSpacing(8)

        btn_file = QPushButton("ファイル選択")
        btn_file.setObjectName("file_button")
        btn_file.setFixedWidth(150)
        btn_file.setFixedHeight(45)
        btn_file.setToolTip(
            "OCRしたい対象の画像を選択してください。<br>"
            "複数ファイルの選択もできます。"
        )
        btn_file.clicked.connect(self.select_files)
        left_buttons.addWidget(btn_file)

        btn_folder = QPushButton("フォルダ選択")
        btn_folder.setObjectName("folder_button")
        btn_folder.setFixedWidth(150)
        btn_folder.setFixedHeight(45)
        btn_folder.setToolTip(
            "指定したフォルダ内の画像を選択。<br>"
            "サブフォルダ内は対象外です。<br>"
            "フォルダ選択をすると現在の選択はリセットされます。"
        )
        btn_folder.clicked.connect(self.select_folder)
        left_buttons.addWidget(btn_folder)

        btn_output = QPushButton("出力先を選択")
        btn_output.setObjectName("output_button")
        btn_output.setFixedWidth(150)
        btn_output.setFixedHeight(45)
        btn_output.setToolTip(
            "解析結果ファイルの保存先フォルダを指定します。<br>"
            "未指定の場合は元ファイルと同じ場所に保存されます。"
        )
        btn_output.clicked.connect(self.select_output)
        left_buttons.addWidget(btn_output)

        right_area = QVBoxLayout()
        right_area.setSpacing(0)

        self.file_list = QListWidget()
        self.file_list.setFixedHeight(150)
        self.file_list.setMinimumWidth(350)
        self.file_list.setToolTip(
            "現在の解析対象となるファイルの一覧です。<br>"
            "ダブルクリックすると、そのファイルをリストから削除します。"
        )
        self.file_list.itemDoubleClicked.connect(self._remove_file_item)

        right_area.addWidget(self.file_list)

        input_area.addLayout(left_buttons)
        input_area.addLayout(right_area)
        parent_layout.addLayout(input_area)

    def _remove_file_item(self, item):
        row = self.file_list.row(item)
        self.file_list.takeItem(row)

        if 0 <= row < len(self.input_paths):
            del self.input_paths[row]

    def refresh_file_list(self):
        self.file_list.clear()
        for p in self.input_paths:
            self.file_list.addItem(str(p))

    # --------------------------------------------------------
    # 3-3-2. 中部 UI（設定項目）
    # --------------------------------------------------------
    def _build_middle_section(self, parent_layout):
        middle_widget = QWidget()
        middle_widget.setMaximumHeight(110)
        middle_widget.setContentsMargins(0, 0, 0, 0)

        grid = QGridLayout(middle_widget)
        grid.setContentsMargins(0, 0, 0, 0)
        grid.setHorizontalSpacing(0)
        grid.setVerticalSpacing(0)

        self._build_middle_contents(grid)

        parent_layout.addWidget(middle_widget, alignment=Qt.AlignHCenter)

    def _build_middle_contents(self, grid):
        """
        中部設定ブロック（3行×4列）
        列の意味:
        1列目: 入力設定（DPI・書字方向）
        2列目: 出力設定・ページ指定
        3-4列目: 図・表の解析設定・保存先
        """
        # 1 行目：DPI / 出力形式 / 図抽出 / 表抽出

        # ▼▼▼ DPI(px) ▼▼▼
        dpi_layout = QHBoxLayout()
        dpi_layout.setContentsMargins(0, 0, 0, 0)
        dpi_layout.setSpacing(4)

        dpi_label = QLabel("DPI(px)：")
        dpi_layout.addWidget(dpi_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)

        dpi_layout.addStretch(1)

        self.dpi_box = QComboBox()
        self.dpi_box.setEditable(True)
        self.dpi_box.addItems(["100", "200", "400", "600"])
        self.dpi_box.setCurrentText("200")
        self.dpi_box.lineEdit().setValidator(QIntValidator(1, 2000))
        self.dpi_box.setFixedWidth(70)
        self.dpi_box.setFixedHeight(28)
        self.dpi_box.setToolTip(
            "PDF を読み込む際の解像度です。<br>"
            "値を上げると精度は向上しますが、処理が重くなります。"
        )

        dpi_wrap = QWidget()
        dpi_wrap.setContentsMargins(0, 0, 10, 0)  # ★ 右マージン 10px
        dpi_wrap_layout = QHBoxLayout(dpi_wrap)
        dpi_wrap_layout.setContentsMargins(0, 0, 0, 0)
        dpi_wrap_layout.addWidget(self.dpi_box)

        dpi_layout.addWidget(dpi_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        dpi_widget = QWidget()
        dpi_widget.setLayout(dpi_layout)
        grid.addWidget(dpi_widget, 0, 0, alignment=Qt.AlignVCenter)


        # ▼▼▼ 出力形式 ▼▼▼
        fmt_layout = QHBoxLayout()
        fmt_layout.setContentsMargins(0, 0, 0, 0)
        fmt_layout.setSpacing(4)

        fmt_label = QLabel("出力形式：")
        fmt_layout.addWidget(fmt_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)

        fmt_layout.addStretch(1)

        self.format_box = QComboBox()
        self.format_box.addItems(["html", "md", "json", "csv", "pdf"])
        self.format_box.setCurrentText("pdf")
        self.format_box.setFixedWidth(80)
        self.format_box.setFixedHeight(28)
        self.format_box.setToolTip("OCR結果の保存形式を選択")

        fmt_wrap = QWidget()
        fmt_wrap.setContentsMargins(0, 0, 10, 0)  # ★ 右マージン 10px
        fmt_wrap_layout = QHBoxLayout(fmt_wrap)
        fmt_wrap_layout.setContentsMargins(0, 0, 0, 0)
        fmt_wrap_layout.addWidget(self.format_box)

        fmt_layout.addWidget(fmt_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        fmt_widget = QWidget()
        fmt_widget.setLayout(fmt_layout)
        grid.addWidget(fmt_widget, 0, 1, alignment=Qt.AlignVCenter)

        # ▼▼▼ 図を抽出する（ラベル + スイッチ） ▼▼▼
        fig_layout = QHBoxLayout()
        fig_layout.setContentsMargins(0, 0, 10, 0)
        fig_layout.setSpacing(4)

        fig_label = QLabel("図を抽出する：")
        fig_layout.addWidget(fig_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        fig_layout.addStretch(1)

        self.figure_check = SwitchWidget()
        self.figure_check.setToolTip(
            "画像内の図形・イラストを検出し、個別の画像として保存します。"
        )

        fig_layout.addWidget(self.figure_check, alignment=Qt.AlignRight | Qt.AlignVCenter)
        fig_widget = QWidget()
        fig_widget.setLayout(fig_layout)
        grid.addWidget(fig_widget, 0, 2, alignment=Qt.AlignVCenter)

        # ▼▼▼ 表を抽出する（ラベル + スイッチ） ▼▼▼
        tbl_layout = QHBoxLayout()
        tbl_layout.setContentsMargins(0, 0, 10, 0)
        tbl_layout.setSpacing(4)

        tbl_label = QLabel("表を抽出する：")
        tbl_layout.addWidget(tbl_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)

        tbl_layout.addStretch(1)

        self.table_check = SwitchWidget()
        self.table_check.setToolTip(
            "画像内の表を検出し、テキストとして構造化された表データに変換します。"
        )
        tbl_layout.addWidget(self.table_check, alignment=Qt.AlignRight | Qt.AlignVCenter)

        tbl_widget = QWidget()
        tbl_widget.setLayout(tbl_layout)
        grid.addWidget(tbl_widget, 0, 3, alignment=Qt.AlignVCenter)

        # 2 行目：高速モード / 解析結果 / 図中文字 / 図幅
        # ▼▼▼ 2 行目：高速モード ▼▼▼
        lite_layout = QHBoxLayout()
        lite_layout.setContentsMargins(0, 0, 0, 0)
        lite_layout.setSpacing(4)

        lite_label = QLabel("高速モード：")
        lite_layout.addWidget(lite_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)

        lite_layout.addStretch(1)

        self.lite_check = SwitchWidget()
        self.lite_check.setContentsMargins(0, 0, 10, 0)
        self.lite_check.setToolTip(
            "一部の解析処理を簡略化し、処理速度を優先します。"
        )

        # ラップして右マージンを確実に反映
        lite_wrap = QWidget()
        lite_wrap.setContentsMargins(0, 0, 10, 0)
        lite_wrap_layout = QHBoxLayout(lite_wrap)
        lite_wrap_layout.setContentsMargins(0, 0, 0, 0)
        lite_wrap_layout.addWidget(self.lite_check)

        lite_layout.addWidget(lite_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        lite_widget = QWidget()
        lite_widget.setLayout(lite_layout)
        grid.addWidget(lite_widget, 1, 0, alignment=Qt.AlignVCenter)


        # ▼▼▼ 解析結果を出力 ▼▼▼
        vis_layout = QHBoxLayout()
        vis_layout.setContentsMargins(0, 0, 0, 0)
        vis_layout.setSpacing(4)

        vis_label = QLabel("解析結果を出力：")
        vis_layout.addWidget(vis_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)

        vis_layout.addStretch(1)

        self.vis_check = SwitchWidget()
        self.vis_check.setContentsMargins(0, 0, 10, 0)
        self.vis_check.setToolTip(
            "解析時の検出結果を可視化した補助的な画像を出力します。"
        )

        vis_wrap = QWidget()
        vis_wrap.setContentsMargins(0, 0, 10, 0)
        vis_wrap_layout = QHBoxLayout(vis_wrap)
        vis_wrap_layout.setContentsMargins(0, 0, 0, 0)
        vis_wrap_layout.addWidget(self.vis_check)

        vis_layout.addWidget(vis_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        vis_widget = QWidget()
        vis_widget.setLayout(vis_layout)
        grid.addWidget(vis_widget, 1, 1, alignment=Qt.AlignVCenter)


        # ▼▼▼ 図の中の文字を抽出 ▼▼▼
        fig_letter_layout = QHBoxLayout()
        fig_letter_layout.setContentsMargins(0, 0, 0, 0)
        fig_letter_layout.setSpacing(4)

        fig_letter_label = QLabel("図の中の文字を抽出：")
        fig_letter_layout.addWidget(fig_letter_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        fig_letter_layout.addStretch(1)

        self.figure_letter_check = SwitchWidget()
        self.figure_letter_check.setContentsMargins(0, 0, 10, 0)
        self.figure_letter_check.setToolTip(
            "図やイラスト内の文字を抽出します。"
        )

        fig_letter_wrap = QWidget()
        fig_letter_wrap.setContentsMargins(0, 0, 10, 0)
        fig_letter_wrap_layout = QHBoxLayout(fig_letter_wrap)
        fig_letter_wrap_layout.setContentsMargins(0, 0, 0, 0)
        fig_letter_wrap_layout.addWidget(self.figure_letter_check)

        fig_letter_layout.addWidget(fig_letter_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        fig_letter_widget = QWidget()
        fig_letter_widget.setLayout(fig_letter_layout)
        grid.addWidget(fig_letter_widget, 1, 2, alignment=Qt.AlignVCenter)

        # ▼▼▼ 図・表の幅(px)▼▼▼
        width_layout = QHBoxLayout()
        width_layout.setContentsMargins(0, 0, 0, 0)
        width_layout.setSpacing(0)

        width_layout.addWidget(QLabel("図・表の幅(px)："))

        self.figure_width_input = QLineEdit()
        self.figure_width_input.setValidator(QIntValidator(1, 5000))
        self.figure_width_input.setFixedWidth(100)
        self.figure_width_input.setFixedHeight(28)
        self.figure_width_input.setToolTip(
            "抽出した図・表を出力する際の表示幅を指定します。<br>"
            "※HTML/Markdown の画像幅指定に使用されます。"
        )
        width_layout.addWidget(self.figure_width_input)

        width_widget = QWidget()
        width_widget.setLayout(width_layout)
        grid.addWidget(width_widget, 1, 3, alignment=Qt.AlignLeft | Qt.AlignVCenter)

        # 3 行目：書字方向 / ページ指定 / 保存先
        direction_layout = QHBoxLayout()
        direction_layout.setContentsMargins(0, 0, 0, 0)
        direction_layout.setSpacing(0)

        direction_label = QLabel("書字方向:")
        direction_layout.addWidget(direction_label, alignment=Qt.AlignVCenter)

        self.direction_box = QComboBox()
        self.direction_box.addItems([
            "自動",
            "横書き",
            "縦書き:上→下",
            "縦書き:右→左"
        ])
        self.direction_box.setFixedWidth(130)
        self.direction_box.setFixedHeight(28)
        self.direction_box.setToolTip(
            "画像内テキストの書字方向を指定します。<br>"
            "自動を選ぶと内容に応じて判定されます。"
        )
        direction_layout.addWidget(self.direction_box, alignment=Qt.AlignVCenter)

        direction_widget = QWidget()
        direction_widget.setLayout(direction_layout)
        grid.addWidget(direction_widget, 2, 0, alignment=Qt.AlignLeft | Qt.AlignVCenter)

        page_layout = QHBoxLayout()
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(0)

        page_layout.addWidget(QLabel("ページ指定:"))

        self.pages_input = QLineEdit()
        self.pages_input.setFixedWidth(120)
        self.pages_input.setFixedHeight(28)
        self.pages_input.setToolTip(
            "PDFの解析対象のページを指定します。<br>"
            "例）1Pと3P～5P → 1,3-5、0 と空欄は全ページ"
        )
        page_layout.addWidget(self.pages_input)

        page_widget = QWidget()
        page_widget.setLayout(page_layout)
        grid.addWidget(page_widget, 2, 1, alignment=Qt.AlignLeft | Qt.AlignVCenter)

        save_layout = QHBoxLayout()
        save_layout.setContentsMargins(0, 0, 0, 0)
        save_layout.setSpacing(0)

        save_layout.addWidget(QLabel("図・表の保存先:"))

        self.figure_dir_input = QLineEdit()
        self.figure_dir_input.setFixedWidth(235)
        self.figure_dir_input.setFixedHeight(28)
        self.figure_dir_input.setToolTip(
            "抽出した図・表の保存先フォルダを指定します。<br>"
            "未指定の場合は出力先フォルダが使用されます。"
        )
        save_layout.addWidget(self.figure_dir_input)

        btn_fig_dir = QPushButton("選択")
        btn_fig_dir.setFixedWidth(60)
        btn_fig_dir.setToolTip(
            "図・表の保存先フォルダを選択します。"
        )
        btn_fig_dir.clicked.connect(self.select_figure_dir)
        save_layout.addWidget(btn_fig_dir)

        save_widget = QWidget()
        save_widget.setLayout(save_layout)

        grid.addWidget(save_widget, 2, 2, 1, 2, alignment=Qt.AlignLeft | Qt.AlignVCenter)

    # --------------------------------------------------------
    # 3-3-3. 下部 UI（実行ボタン・ログ）
    # --------------------------------------------------------
    def _build_bottom_section(self, parent_layout):
        """
        下部エリア:
        - 実行ボタン
        - ログ表示
        """
        btn_run = QPushButton("実行")
        btn_run.setObjectName("run_button")
        btn_run.setFixedHeight(45)
        btn_run.clicked.connect(self.run_yomitoku)
        btn_run.setStyleSheet(
            "QPushButton {\n"
            "        font-size: 20px;\n"
            "        color: black;\n"
            "        background-color: #e0e0e0;\n"
            "        border: 1px solid #888;\n"
            "        border-radius: 6px;\n"
            "}"
        )
        parent_layout.addWidget(btn_run)
        self.log_view = QTextEdit()
        self.log_view.setReadOnly(True)
        self.log_view.setFixedHeight(300)
        self.log_view.setToolTip(
            "処理の進行状況やエラー内容を表示します。"
        )
        parent_layout.addWidget(self.log_view)
    # --------------------------------------------------------
    # 3-4. UI 有効/無効・進捗
    # --------------------------------------------------------
    def disable_ui(self):
        self.findChild(QPushButton, "run_button").setEnabled(False)
        self.findChild(QPushButton, "output_button").setEnabled(False)
        self.findChild(QPushButton, "file_button").setEnabled(False)
        self.findChild(QPushButton, "folder_button").setEnabled(False)

    def enable_ui(self):
        self.findChild(QPushButton, "run_button").setEnabled(True)
        self.findChild(QPushButton, "output_button").setEnabled(True)
        self.findChild(QPushButton, "file_button").setEnabled(True)
        self.findChild(QPushButton, "folder_button").setEnabled(True)

    def update_run_button(self, value: int):
        btn = self.findChild(QPushButton, "run_button")

        if value >= self.total_files:
            btn.setText("完了")
            btn.setStyleSheet(
                "QPushButton {"
                "    font-size: 20px;"
                "    font-weight: bold;"
                "    color: black;"
                "    background-color: #00C896;"
                "    border: 1px solid #888;"
                "    border-radius: 6px;"
                "}"
            )
            return

        ratio = value / max(1, self.total_files)
        stop = f"{ratio:.2f}"

        btn.setStyleSheet(
            f"QPushButton {{"
            f"    font-size: 20px;"
            f"    font-weight: bold;"
            f"    color: black;"
            f"    border: 1px solid #888;"
            f"    border-radius: 6px;"
            f"    background: qlineargradient("
            f"        x1:0, y1:0, x2:1, y2:0,"
            f"        stop:0 rgba(0, 200, 150, 255),"
            f"        stop:{stop} rgba(0, 200, 150, 255),"
            f"        stop:{stop} rgba(224, 224, 224, 255),"
            f"        stop:1 rgba(224, 224, 224, 255)"
            f"    );"
            f"}}"
        )
        btn.setText(f"{value} / {self.total_files}")

    def reset_run_button(self):
        btn = self.findChild(QPushButton, "run_button")
        btn.setText("実行")
        btn.setStyleSheet(
            "QPushButton {"
            "    font-size: 20px;"
            "    color: black;"
            "    background-color: #e0e0e0;"
            "    border: 1px solid #888;"
            "    border-radius: 6px;"
            "}"
        )

    # --------------------------------------------------------
    # 3-5. Drag & Drop
    # --------------------------------------------------------
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
        else:
            event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if not urls:
            return

        for url in urls:
            path = Path(url.toLocalFile())
            if path.exists():
                self.input_paths.append(path)
                self.file_list.addItem(str(path))

        self.log(f"{len(urls)} 件をドロップで追加しました")
        self.reset_run_button()

    # --------------------------------------------------------
    # 3-6. 入力・出力選択
    # --------------------------------------------------------
    def select_files(self):
        # 最後に開いたフォルダを取得
        last_dir = self.config.get("Settings", "last_file_dir", fallback="")

        patterns = " ".join(f"*{ext}" for ext in self.SUPPORTED_EXT)
        name_filter = f"対応ファイル ({patterns})"

        files, _ = QFileDialog.getOpenFileNames(
            self,
            "ファイルを選択",
            last_dir,   # ★ ここが重要
            name_filter
        )

        if not files:
            return

        # ★ 最後に開いたフォルダを保存
        folder = str(Path(files[0]).parent)
        self.config["Settings"]["last_file_dir"] = folder
        self.save_config()

        # 選択されたファイルを UI に反映
        self.file_list.clear()
        self.input_paths = []
        for f in files:
            path = Path(f)
            if path.suffix.lower() in self.SUPPORTED_EXT:
                self.file_list.addItem(str(path))
                self.input_paths.append(path)

        self.reset_run_button()

    def select_folder(self):
        last_dir = self.config.get("Settings", "last_folder_dir", fallback="")

        folder = QFileDialog.getExistingDirectory(self, "フォルダを選択", last_dir)
        if not folder:
            return

        # ★ 最後に選択したフォルダを保存
        self.config["Settings"]["last_folder_dir"] = folder
        self.save_config()

        self.file_list.clear()
        self.input_paths = []

        path = Path(folder)
        files = sorted(
            f for f in path.iterdir()
            if f.is_file() and f.suffix.lower() in self.SUPPORTED_EXT
        )

        for f in files:
            self.file_list.addItem(str(f))
            self.input_paths.append(f)

        self.reset_run_button()

    def select_output(self):
        folder = QFileDialog.getExistingDirectory(self, "出力先フォルダ")
        if folder:
            self.output_dir = Path(folder)
            if "Settings" not in self.config:
                self.config["Settings"] = {}
            self.config["Settings"]["output_dir"] = str(self.output_dir)
            self.save_config()

    def select_figure_dir(self):
        folder = QFileDialog.getExistingDirectory(self, "図の保存先フォルダ")
        if folder:
            self.figure_dir_input.setText(folder)

    # --------------------------------------------------------
    # 3-7. 実行処理（スレッド化）
    # --------------------------------------------------------
    def run_yomitoku(self):
        if not self.input_paths:
            self.log("入力が選択されていません。")
            return

        self.total_files = len(self.input_paths)

        import os
        exe_path = self.yomitoku_path
        if exe_path is None:
            self.log("YomiToku のパスが設定されていません。")
            return

        fmt = self.format_box.currentText()

        if getattr(self, "output_dir", None) is None:
            first = self.input_paths[0]
            outdir_path = first.parent if first.is_file() else first
        else:
            outdir_path = self.output_dir

        dpi_text = self.dpi_box.currentText().strip()
        if not dpi_text.isdigit():
            self.log("DPI が不正なため 200 にリセットしました。")
            dpi_text = "200"
            self.dpi_box.setCurrentText("200")

        dpi_value = int(dpi_text)
        if dpi_value <= 0 or dpi_value > 2000:
            self.log("エラー: DPI の値が対応範囲外です。")
            return

        dpi = str(dpi_value)

        ro_map = {
            "自動": "auto",
            "横書き": "left2right",
            "縦書き:上→下": "top2bottom",
            "縦書き:右→左": "right2left"
        }
        reading_order = ro_map[self.direction_box.currentText()]

        pages_raw = self.pages_input.text().strip()
        if pages_raw == "" or pages_raw == "0":
            pages = ""
        else:
            import re
            pattern = r"^(\d+(-\d+)?)(,(\d+(-\d+)?))*$"
            if not re.match(pattern, pages_raw):
                self.log("ページ指定の形式が不正です。例：1,3-5")
                return
            pages = pages_raw

        figure = self.figure_check.isChecked()
        table = self.table_check.isChecked()

        lite = self.lite_check.isChecked()
        vis = self.vis_check.isChecked()

        self.disable_ui()
        self.reset_run_button()

        self.thread = QThread()
        self.worker = YomiTokuWorker(
            exe_path,
            self.input_paths,
            outdir_path,
            fmt,
            dpi,
            reading_order,
            pages,
            figure,
            table,
            lite,
            vis
        )
        self.worker.moveToThread(self.thread)

        self.thread.started.connect(self.worker.run)
        self.worker.log_signal.connect(self.log)
        self.worker.progress_signal.connect(self.update_run_button)
        self.worker.finished.connect(self.thread.quit)
        self.worker.finished.connect(self.worker.deleteLater)
        self.thread.finished.connect(self.thread.deleteLater)
        self.worker.finished.connect(self.enable_ui)

        self.thread.start()
        self.log("=== スレッド開始 ===")

# ============================================================
# 4. Main
# ============================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)

    # ツールチップの挙動（即表示・無制限）
    class TooltipStyle(QProxyStyle):
        def styleHint(self, hint, option=None, widget=None, returnData=None):
            if hint == QStyle.SH_ToolTip_WakeUpDelay:
                return 100
            if hint == QStyle.SH_ToolTip_FallAsleepDelay:
                return 0
            return super().styleHint(hint, option, widget, returnData)

    app.setStyle(TooltipStyle())

    # ツールチップの見た目（背景色・文字色・幅など）
    app.setStyleSheet("""
        QToolTip {
            background-color: #333333;
            color: #ffffff;
            border: 1px solid #aaaaaa;
            padding: 6px;
            font-size: 12pt;
            max-width: 1000px;
            white-space: nowrap;
        }
    """)

    gui = YomiTokuGUI()
    gui.show()
    sys.exit(app.exec())
