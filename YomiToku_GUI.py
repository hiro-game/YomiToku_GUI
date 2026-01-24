# ============================================================
# 1. import 群
# ============================================================
import os
import subprocess
import configparser
import importlib
import sys

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

# ============================================================
# 2. クラス：SwitchWidget（ON/OFF スイッチ UI）
# ============================================================
class SwitchWidget(QWidget):

    # ON/OFF が切り替わったときに通知するシグナル
    toggled = Signal(bool)

    # 2-1. 初期化（スイッチの基本状態を設定）
    # --------------------------------------------------------
    def __init__(self, parent=None):
        super().__init__(parent)
        self._checked = False
        self.setFixedSize(42, 20)

    # 2-2. 推奨サイズの返却
    # --------------------------------------------------------
    def sizeHint(self):
        return QSize(42, 20)

    # 2-3. チェック状態の取得
    # --------------------------------------------------------
    def isChecked(self):
        return self._checked

    # 2-4. チェック状態の設定
    # --------------------------------------------------------
    def setChecked(self, value: bool):
        value = bool(value)
        if self._checked != value:
            self._checked = value
            self.toggled.emit(self._checked)
            self.update()

    # 2-5. マウスクリックで ON/OFF 切り替え
    # --------------------------------------------------------
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self._checked = not self._checked
            self.toggled.emit(self._checked)
            self.update()

    # 2-6. スイッチ描画処理
    # --------------------------------------------------------
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)

        # 背景色（ON/OFF で色を変える）
        bg_color = QColor(160, 160, 160) if not self._checked else QColor(210, 210, 210)
        painter.setBrush(QBrush(bg_color))
        painter.setPen(Qt.NoPen)
        painter.drawRoundedRect(self.rect(), 4, 4)

        # ハンドル描画
        handle_size = 16
        y = (self.height() - handle_size) // 2
        x = self.width() - handle_size - 2 if self._checked else 2
        handle_rect = QRect(x, y, handle_size, handle_size)

        painter.setBrush(QBrush(QColor(240, 240, 240)))
        painter.setPen(QPen(QColor(150, 150, 150), 1))
        painter.drawRoundedRect(handle_rect, 4, 4)

# ============================================================
# 3. 関数：detect_device（PyTorch のデバイス自動判定）
# ============================================================
def detect_device():
    """
    初回起動時のみ呼ばれる。
    PyTorch を import し、利用可能なデバイスを判定して返す。
    2 回目以降は ini の値を使うため、この関数は呼ばれない。
    """

    try:
        import torch
    except ImportError:
        return "cpu"

    # CUDA
    if torch.cuda.is_available():
        return "cuda"

    # 将来の NPU / iGPU 対応（仮）
    if hasattr(torch, "npu") and torch.npu.is_available():
        return "npu"

    if hasattr(torch, "xpu") and torch.xpu.is_available():
        return "xpu"

    return "cpu"

# ============================================================
# 2. Worker クラス（バックエンド処理）
# ============================================================
class YomiTokuWorker(QObject):

    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished = Signal()

    # 2-1. 初期化（パラメータ保存のみ）
    # --------------------------------------------------------
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
        figure_letter,
        figure_width,
        combine,
        ignore_line_break,
        ignore_meta,
        encoding,
        lite,
        vis,
        device,
        font_path,
        td_cfg,
        tr_cfg,
        lp_cfg,
        tsr_cfg,
        figure_dir
    ):
        super().__init__()

        self.exe_path = exe_path
        self.input_paths = input_paths
        self.outdir_path = outdir_path
        self.fmt = fmt
        self.dpi = dpi
        self.reading_order = reading_order
        self.pages = pages
        self.figure = figure
        self.figure_letter = figure_letter
        self.figure_width = figure_width
        self.figure_dir = figure_dir
        self.combine = combine
        self.ignore_line_break = ignore_line_break
        self.ignore_meta = ignore_meta
        self.encoding = encoding
        self.lite = lite
        self.vis = vis
        self.device = device
        self.font_path = font_path
        self.td_cfg = td_cfg
        self.tr_cfg = tr_cfg
        self.lp_cfg = lp_cfg
        self.tsr_cfg = tsr_cfg

    # 2-2. 実行処理（run に移動）
    # --------------------------------------------------------
    def run(self):

        total = len(self.input_paths)
        self.log_signal.emit(f"----- {total} 件の処理を開始 -----")

        for idx, input_path in enumerate(self.input_paths, start=1):

            progress = int((idx - 1) / total * 100)
            self.progress_signal.emit(progress)

            self.log_signal.emit(f"[{idx}/{total}] 処理中: {input_path}")

            # ★ 実行コマンドの構築
            cmd = [
                str(self.exe_path),
                str(input_path),
                "-o", str(self.outdir_path),
                "-f", self.fmt,
                "-d", self.device
            ]

            # ★ オプション追加
            if self.dpi:
                cmd += ["--dpi", self.dpi]

            if self.reading_order:
                cmd += ["--reading_order", self.reading_order]

            if self.pages:
                cmd += ["--pages", self.pages]

            if self.figure:
                cmd.append("--figure")

            if self.figure_letter:
                cmd.append("--figure_letter")

            if self.figure_width:
                cmd += ["--figure_width", self.figure_width]

            if self.figure_dir:
                cmd += ["--figure_dir", self.figure_dir]

            if self.combine:
                cmd.append("--combine")

            if self.ignore_line_break:
                cmd.append("--ignore_line_break")

            if self.ignore_meta:
                cmd.append("--ignore_meta")

            if self.encoding:
                cmd += ["--encoding", self.encoding]

            if self.lite:
                cmd.append("--lite")

            if self.vis:
                cmd.append("--vis")

            if self.td_cfg:
                cmd += ["--td_cfg", self.td_cfg]

            if self.tr_cfg:
                cmd += ["--tr_cfg", self.tr_cfg]

            if self.lp_cfg:
                cmd += ["--lp_cfg", self.lp_cfg]

            if self.tsr_cfg:
                cmd += ["--tsr_cfg", self.tsr_cfg]

            if self.font_path:
                cmd += ["--font_path", self.font_path]

            # ★ 最終コマンドをログ出力
            self.log_signal.emit("実行コマンド: " + " ".join(cmd))

            # ★ サブプロセス実行
            process = subprocess.Popen(
                cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True
            )

            for line in process.stdout:
                self.log_signal.emit(line.rstrip())

            process.wait()
            self.log_signal.emit(f"完了: {input_path} (終了コード: {process.returncode})")

        self.progress_signal.emit(100)
        self.log_signal.emit("----- 全ての処理が完了しました -----")
        self.finished.emit()

# ============================================================
# 3. GUI 本体
# ============================================================
class YomiTokuGUI(QWidget):

    # ★ 対応拡張子（ドラッグ＆ドロップやファイル選択で使用）
    SUPPORTED_EXT = [".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"]

    # --------------------------------------------------------
    # 3-1. 初期化（GUI 全体の初期セットアップ）
    # --------------------------------------------------------
    def __init__(self):
        super().__init__()

        # ★ GUI 全体のフォント設定
        font = QFont()
        font.setPointSize(12)
        self.setFont(font)

        # ★ ini_path と config を最初に作る（重要）
        #   - load_all_settings() で ini を読むため、必ず最初に作成する
        self.ini_path = os.path.join(os.path.dirname(__file__), "YomiToku_GUI.ini")
        self.config = configparser.ConfigParser()

        # ★ 内部状態の初期化（フラグや変数の初期値）
        self._init_basic_state()

        # ★ UI 構築（ウィジェット・レイアウトの生成）
        self._build_ui()

        # ★ 初期設定の読み込み（UI の初期値を設定）
        self._load_initial_config()

        # ★ 設定を UI に反映（ini_path があるので安全）
        #   - save_settings / save_log / yomitoku_path / device などを読み込む
        self.load_all_settings()

        # ★ YomiToku のパスを読み込む（初回は自動検出）
        self.load_yomitoku_path()

        # ★ ウィンドウサイズを内容に合わせて固定
        self.adjustSize()
        self.setFixedSize(self.size())

        # ★ 起動ログ
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log(f"=== App started at {timestamp} ===")

    # --------------------------------------------------------
    # 3-2. YomiToku のパスを読み込む（初回は自動取得）
    #    - 設定ファイルに有効なパスがあればそれを最優先
    #    - 無ければ自動検出（site → where/which → ~/.local/bin）
    # --------------------------------------------------------
    def load_yomitoku_path(self):
        cfg = self.config

        # ★ Settings セクションが無い場合は作成
        if "Settings" not in cfg:
            cfg["Settings"] = {}

        settings = cfg["Settings"]

        # --------------------------------------------------------
        # 1. 設定ファイルにパスがある場合はそれを最優先（ユーザーの意思を尊重）
        # --------------------------------------------------------
        if "yomitoku_path" in settings:
            raw = settings["yomitoku_path"].strip()

            if raw:
                path = Path(raw)

                # ★ 有効なパスなら即採用（自動検出は行わない）
                if path.exists():
                    self.yomitoku_path = path
                    return

                # パスが存在しない場合のみ自動検出へ
                self.log(f"設定ファイルの YomiToku パスが存在しません: {raw}")

            else:
                self.log("設定ファイルに yomitoku_path が空で保存されています。")

        else:
            self.log("設定ファイルに yomitoku_path がありません。")

        # --------------------------------------------------------
        # 2. 自動検出（site → Scripts/bin → where/which）
        # --------------------------------------------------------
        auto_path = self.find_yomitoku_exe()

        if auto_path and auto_path.exists():
            self.yomitoku_path = auto_path

            # ★ 設定にまだ値が無い場合のみ保存（既存設定は上書きしない）
            if not settings.get("yomitoku_path", "").strip():
                settings["yomitoku_path"] = str(auto_path)
                self.save_config()
                self.log(f"YomiToku パスを自動検出し、設定に保存しました: {auto_path}")
            else:
                self.log(f"YomiToku パスを自動検出しました（設定は既存値を維持）: {auto_path}")

            return

        # --------------------------------------------------------
        # 3. 自動検出できなかった場合
        # --------------------------------------------------------
        self.yomitoku_path = None
        self.log("YomiToku のパスを自動検出できませんでした。設定画面から手動で指定してください。")

    # --------------------------------------------------------
    # 3-4. YomiToku 実行ファイルの自動検出
    #    - site.getsitepackages / getusersitepackages を利用
    #    - Scripts / bin を総当たり
    #    - where / which をフォールバックとして使用
    #    - 複数見つかった場合は「より新しい Python バージョン」を優先
    # --------------------------------------------------------
    def find_yomitoku_exe(self):

        candidates = []

        # ★ 候補リストに追加（重複排除）
        def add_candidate(p: Path):
            if p and p.exists():
                p = p.resolve()
                if p not in candidates:
                    candidates.append(p)

        # --------------------------------------------------------
        # 1. site.getsitepackages / getusersitepackages → Scripts/bin を探索
        # --------------------------------------------------------
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

        # --------------------------------------------------------
        # 2. where / which をフォールバックとして使用
        # --------------------------------------------------------
        exe_name = "yomitoku.exe" if sys.platform.startswith("win") else "yomitoku"
        exe = shutil.which(exe_name)
        if exe:
            add_candidate(Path(exe))

        # --------------------------------------------------------
        # 3. Unix 系の ~/.local/bin/yomitoku もチェック
        # --------------------------------------------------------
        if not sys.platform.startswith("win"):
            local_bin = Path.home() / ".local" / "bin" / "yomitoku"
            if local_bin.exists():
                add_candidate(local_bin)

        # 候補が無い場合
        if not candidates:
            return None

        # --------------------------------------------------------
        # 4. 複数見つかった場合は「Python バージョンが新しいもの」を優先
        # --------------------------------------------------------
        def version_key(p: Path):
            s = str(p)

            # 例: Python313, python311 など
            m = re.search(r"[Pp]ython(?:3)?(\d)(\d)", s)
            if m:
                return (int(m.group(1)), int(m.group(2)))

            # 例: python3.11, Python3.10 など
            m = re.search(r"[Pp]ython(\d)\.(\d+)", s)
            if m:
                return (int(m.group(1)), int(m.group(2)))

            # バージョン情報が取れない場合は最低優先
            return (0, 0)

        candidates.sort(key=version_key, reverse=True)
        return candidates[0]

    # --------------------------------------------------------
    # 3-5. ログ表示（ログビューに追記）
    # --------------------------------------------------------
    def log(self, text):
        self.log_view.append(text)

    # --------------------------------------------------------
    # 3-6. 状態・設定関連（初期化）
    # --------------------------------------------------------
    def _init_basic_state(self):
        self.setWindowTitle("YomiToku_GUI")
        self.resize(900, 650)
        self.setAcceptDrops(True)

        # ★ 設定ファイルのパス
        self.config_path = Path("YomiToku_GUI.ini")
        self.config = configparser.ConfigParser()

        # ★ 入力ファイル・出力ディレクトリの初期状態
        self.input_paths = []
        self.output_dir = None

    # --------------------------------------------------------
    # 3-7. 初期設定の読み込み（UI 初期値の反映）
    # --------------------------------------------------------
    def _load_initial_config(self):
        self.load_config()

        # 出力先ディレクトリ
        if "Settings" in self.config and "output_dir" in self.config["Settings"]:
            self.output_dir = Path(self.config["Settings"]["output_dir"])

        # 保存フラグ
        self.save_settings_flag = self.config.get("Settings", "save_settings", fallback="1") == "1"
        self.save_log_flag = self.config.get("Settings", "save_log", fallback="0") == "1"

    # --------------------------------------------------------
    # 3-8. 設定ファイルの読み込み
    # --------------------------------------------------------
    def load_config(self):
        if self.config_path.exists():
            self.config.read(self.config_path, encoding="utf-8")
        else:
            # ★ 初回起動時のデフォルト設定
            self.config["Settings"] = {
                 "save_settings": "0",
                 "save_log": "0"
            }
            self.save_config()

    # --------------------------------------------------------
    # 3-9. 設定ファイルの保存
    # --------------------------------------------------------
    def save_config(self):
        with open(self.config_path, "w", encoding="utf-8") as f:
            self.config.write(f)

    # --------------------------------------------------------
    # 3-10. UI の設定値をすべて保存
    # --------------------------------------------------------
    def save_all_settings(self):
        cfg = self.config

        if "Settings" not in cfg:
            cfg["Settings"] = {}

        s = cfg["Settings"]

        # 保存フラグ
        s["save_settings"] = "1" if self.save_settings_flag else "0"
        s["save_log"] = "1" if self.save_log_flag else "0"

        # 中部の設定内容
        s["format"] = str(self.format_box.currentIndex())
        s["reading_order"] = str(self.direction_box.currentIndex())
        s["dpi"] = self.dpi_box.currentText()
        s["pages"] = self.pages_input.text()
        s["figure_width"] = self.figure_width_input.text()
        s["figure_dir"] = self.figure_dir_input.text()

        # チェックボックス類
        s["figure"] = "1" if self.figure_check.isChecked() else "0"
        s["table"] = "1" if self.table_check.isChecked() else "0"
        s["lite"] = "1" if self.lite_check.isChecked() else "0"
        s["vis"] = "1" if self.vis_check.isChecked() else "0"
        s["figure_letter"] = "1" if self.figure_letter_check.isChecked() else "0"

        self.save_config()

    # --------------------------------------------------------
    # 3-11. 設定ファイルの内容を UI に反映
    # --------------------------------------------------------
    def load_all_settings(self):
        cfg = self.config
        if "Settings" not in cfg:
            return

        s = cfg["Settings"]

        # ★ device が無い、または空なら探査して追記（初回のみ）
        if "device" not in s or not s["device"]:
            device = detect_device()
            s["device"] = device
            self.config.set("Settings", "device", device)

            # ★ ini_path に保存（初回のみ）
            with open(self.ini_path, "w", encoding="utf-8") as f:
                self.config.write(f)
        else:
            device = s["device"]
        self.device = device

        # ★ save_settings=1 のときだけ last_* と output_dir を読み込む
        if s.get("save_settings", "0") == "1":
            self.last_file_dir = s.get("last_file_dir", "")
            self.last_folder_dir = s.get("last_folder_dir", "")
            self.output_dir = Path(s.get("output_dir", "")) if s.get("output_dir") else None
        else:
            self.last_file_dir = ""
            self.last_folder_dir = ""
            self.output_dir = None

        # 中部の設定内容
        if "format" in s:
            self.format_box.setCurrentIndex(int(s["format"]))

        if "reading_order" in s:
            self.direction_box.setCurrentIndex(int(s["reading_order"]))

        if "dpi" in s:
            self.dpi_box.setCurrentText(s["dpi"])

        if "pages" in s:
            self.page_input.setText(s["pages"])

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

        # ★ vis_check は GUI に無いので削除（ini の vis は run_yomitoku() で処理）
        # if "vis" in s:
        #     self.vis_check.setChecked(s["vis"] == "1")

        if "figure_letter" in s:
            self.figure_letter_check.setChecked(s["figure_letter"] == "1")

    # --------------------------------------------------------
    # 3-12. ウィンドウ終了時の処理
    # --------------------------------------------------------
    def closeEvent(self, event):

        # 設定保存（save_settings=1 のときのみ）
        if self.save_settings_flag:
            self.save_all_settings()

        # ログ保存（save_log=1 のときのみ）
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
        3-4列目: 図の解析設定・保存先
        """

    # --------------------------------------------------------
    # 3-3-2. 中部 UI（設定項目）
    # --------------------------------------------------------
    def _build_middle_contents(self, grid):
        """
        中部設定ブロック（3行×4列）
        列の意味:
        1列目: 入力設定（DPI・書字方向）
        2列目: 出力設定・ページ指定
        3-4列目: 図の解析設定・保存先
        """

        # --------------------------------------------------------
        # 1 行目：DPI / ヘッダー等を無視 / 高速モード / 図を抽出
        # --------------------------------------------------------
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
            "PDFを読み込む際のDPIを指定、数字を大きくすると精度は上がるが負荷が掛かります。"
        )

        dpi_wrap = QWidget()
        dpi_wrap.setContentsMargins(0, 0, 10, 0)
        dpi_wrap_layout = QHBoxLayout(dpi_wrap)
        dpi_wrap_layout.setContentsMargins(0, 0, 0, 0)
        dpi_wrap_layout.addWidget(self.dpi_box)
        dpi_layout.addWidget(dpi_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        dpi_widget = QWidget()
        dpi_widget.setLayout(dpi_layout)
        grid.addWidget(dpi_widget, 0, 0, alignment=Qt.AlignVCenter)

        # ▼▼▼ ヘッダー等を無視 ▼▼▼
        meta_layout = QHBoxLayout()
        meta_layout.setContentsMargins(0, 0, 10, 0)
        meta_layout.setSpacing(4)

        meta_label = QLabel("ヘッダー等を無視：")
        meta_layout.addWidget(meta_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        meta_layout.addStretch(1)

        self.ignore_meta_check = SwitchWidget()
        self.ignore_meta_check.setToolTip(
            "画像の上下に書き込まれたページ数などのヘッダーとフッターはOCRしません。"
        )
        meta_layout.addWidget(self.ignore_meta_check, alignment=Qt.AlignRight | Qt.AlignVCenter)

        meta_widget = QWidget()
        meta_widget.setLayout(meta_layout)
        grid.addWidget(meta_widget, 0, 1, alignment=Qt.AlignVCenter)

        # ▼▼▼ 高速モード ▼▼▼
        lite_layout = QHBoxLayout()
        lite_layout.setContentsMargins(0, 0, 10, 0)
        lite_layout.setSpacing(4)

        lite_label = QLabel("高速モード：")
        lite_layout.addWidget(lite_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        lite_layout.addStretch(1)

        self.lite_check = SwitchWidget()
        self.lite_check.setToolTip(
            "処理を高速化しますが、一部の解析精度が低下する場合があります。"
        )
        lite_layout.addWidget(self.lite_check, alignment=Qt.AlignRight | Qt.AlignVCenter)

        lite_widget = QWidget()
        lite_widget.setLayout(lite_layout)
        grid.addWidget(lite_widget, 0, 2, alignment=Qt.AlignVCenter)

        # ▼▼▼ 図を抽出する ▼▼▼
        fig_layout = QHBoxLayout()
        fig_layout.setContentsMargins(0, 0, 10, 0)
        fig_layout.setSpacing(4)

        fig_label = QLabel("図を抽出：")
        fig_layout.addWidget(fig_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        fig_layout.addStretch(1)

        self.figure_check = SwitchWidget()
        self.figure_check.setToolTip(
            "画像内の図やイラストを検出して個別に抽出・保存します。"
        )
        fig_layout.addWidget(self.figure_check, alignment=Qt.AlignRight | Qt.AlignVCenter)

        fig_widget = QWidget()
        fig_widget.setLayout(fig_layout)
        grid.addWidget(fig_widget, 0, 3, alignment=Qt.AlignVCenter)

        # --------------------------------------------------------
        # 2 行目：ページ指定 / 書字方向 / 文字コード / 図の中の文字を抽出
        # --------------------------------------------------------
        # ▼▼▼ ページ指定 ▼▼▼
        page_layout = QHBoxLayout()
        page_layout.setContentsMargins(0, 0, 0, 0)
        page_layout.setSpacing(4)

        page_label = QLabel("ページ指定：")
        page_layout.addWidget(page_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        page_layout.addStretch(1)

        page_wrap = QWidget()
        page_wrap.setContentsMargins(0, 0, 10, 0)
        page_wrap_layout = QHBoxLayout(page_wrap)
        page_wrap_layout.setContentsMargins(0, 0, 0, 0)

        self.pages_input = QLineEdit()
        self.pages_input.setPlaceholderText("例：1,3-5")
        self.pages_input.setFixedWidth(120)
        self.pages_input.setFixedHeight(28)
        self.pages_input.setToolTip(
            "PDFから読込むページを指定します。<BR>"
            "1P、3P～5Pを読み込む場合<BR>例：1,3-5。"
            "　(0か空白で全選択)"
        )
        page_wrap_layout.addWidget(self.pages_input)

        page_layout.addWidget(page_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        page_widget = QWidget()
        page_widget.setLayout(page_layout)
        grid.addWidget(page_widget, 1, 0, alignment=Qt.AlignVCenter)

        # ▼▼▼ 書字方向 ▼▼▼
        dir_layout = QHBoxLayout()
        dir_layout.setContentsMargins(0, 0, 0, 0)
        dir_layout.setSpacing(4)

        dir_label = QLabel("書字方向：")
        dir_layout.addWidget(dir_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        dir_layout.addStretch(1)

        dir_wrap = QWidget()
        dir_wrap.setContentsMargins(0, 0, 10, 0)
        dir_wrap_layout = QHBoxLayout(dir_wrap)
        dir_wrap_layout.setContentsMargins(0, 0, 0, 0)

        self.direction_box = QComboBox()
        self.direction_box.addItems([
            "自動",
            "横書き",
            "縦書き:上→下",
            "縦書き:右→左",
        ])
        self.direction_box.setCurrentText("自動")
        self.direction_box.setFixedWidth(120)
        self.direction_box.setFixedHeight(28)
        self.direction_box.setToolTip(
            "画像内テキストの書字方向を指定します。<br>"
            "自動を選ぶと内容に応じて判定されます。"
        )
        dir_wrap_layout.addWidget(self.direction_box)

        dir_layout.addWidget(dir_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        dir_widget = QWidget()
        dir_widget.setLayout(dir_layout)
        grid.addWidget(dir_widget, 1, 1, alignment=Qt.AlignVCenter)

        # ▼▼▼ 文字コード ▼▼▼
        enc_layout = QHBoxLayout()
        enc_layout.setContentsMargins(0, 0, 0, 0)
        enc_layout.setSpacing(4)

        enc_label = QLabel("文字コード：")
        enc_layout.addWidget(enc_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        enc_layout.addStretch(1)

        enc_wrap = QWidget()
        enc_wrap.setContentsMargins(0, 0, 10, 0)
        enc_wrap_layout = QHBoxLayout(enc_wrap)
        enc_wrap_layout.setContentsMargins(0, 0, 0, 0)

        self.encoding_box = QComboBox()
        self.encoding_box.addItems(["utf-8", "shift_jis"])
        self.encoding_box.setCurrentText("utf-8")
        self.encoding_box.setFixedWidth(90)
        self.encoding_box.setFixedHeight(28)
        self.encoding_box.setToolTip("出力テキストの文字コードを指定します。")

        enc_wrap_layout.addWidget(self.encoding_box)

        enc_layout.addWidget(enc_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        enc_widget = QWidget()
        enc_widget.setLayout(enc_layout)
        grid.addWidget(enc_widget, 1, 2, alignment=Qt.AlignVCenter)

        # ▼▼▼ 図の中の文字を抽出 ▼▼▼
        figlet_layout = QHBoxLayout()
        figlet_layout.setContentsMargins(0, 0, 0, 0)
        figlet_layout.setSpacing(4)

        figlet_label = QLabel("図内の文字を抽出：")
        figlet_layout.addWidget(figlet_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        figlet_layout.addStretch(1)

        figlet_wrap = QWidget()
        figlet_wrap.setContentsMargins(0, 0, 10, 0)
        figlet_wrap_layout = QHBoxLayout(figlet_wrap)
        figlet_wrap_layout.setContentsMargins(0, 0, 0, 0)

        self.figure_letter_check = SwitchWidget()
        self.figure_letter_check.setToolTip(
            "図の内部に含まれる文字を OCR で抽出します。"
        )
        figlet_wrap_layout.addWidget(self.figure_letter_check)

        figlet_layout.addWidget(figlet_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        figlet_widget = QWidget()
        figlet_widget.setLayout(figlet_layout)
        grid.addWidget(figlet_widget, 1, 3, alignment=Qt.AlignVCenter)

        # --------------------------------------------------------
        # 3 行目：ページを結合する / 改行を無視 / 出力形式 / 図の幅
        # --------------------------------------------------------
        # ▼▼▼ 複数ページをまとめる ▼▼▼
        combine_layout = QHBoxLayout()
        combine_layout.setContentsMargins(0, 0, 0, 0)
        combine_layout.setSpacing(4)

        combine_label = QLabel("複数ページをまとめる：")
        combine_layout.addWidget(combine_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        combine_layout.addStretch(1)

        combine_wrap = QWidget()
        combine_wrap.setContentsMargins(0, 0, 10, 0)
        combine_wrap_layout = QHBoxLayout(combine_wrap)
        combine_wrap_layout.setContentsMargins(0, 0, 0, 0)

        self.combine_check = SwitchWidget()
        self.combine_check.setToolTip(
            "通常１つのPDFから読み取った複数のページは一つづつのファイルに保存されますが<BR>出力がPDF場合に単一のPDFファイルに結合します。"
        )
        combine_wrap_layout.addWidget(self.combine_check)

        combine_layout.addWidget(combine_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        combine_widget = QWidget()
        combine_widget.setLayout(combine_layout)
        grid.addWidget(combine_widget, 2, 0, alignment=Qt.AlignVCenter)

        # ▼▼▼ 改行を無視 ▼▼▼
        lb_layout = QHBoxLayout()
        lb_layout.setContentsMargins(0, 0, 0, 0)
        lb_layout.setSpacing(4)

        lb_label = QLabel("改行を無視：")
        lb_layout.addWidget(lb_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        lb_layout.addStretch(1)

        lb_wrap = QWidget()
        lb_wrap.setContentsMargins(0, 0, 10, 0)
        lb_wrap_layout = QHBoxLayout(lb_wrap)
        lb_wrap_layout.setContentsMargins(0, 0, 0, 0)

        self.ignore_lb_check = SwitchWidget()
        self.ignore_lb_check.setToolTip(
            "OCR 結果の改行を無視し、連続した文章として扱います。"
        )
        lb_wrap_layout.addWidget(self.ignore_lb_check)

        lb_layout.addWidget(lb_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        lb_widget = QWidget()
        lb_widget.setLayout(lb_layout)
        grid.addWidget(lb_widget, 2, 1, alignment=Qt.AlignVCenter)

        # ▼▼▼ 出力形式 ▼▼▼
        fmt_layout = QHBoxLayout()
        fmt_layout.setContentsMargins(0, 0, 0, 0)
        fmt_layout.setSpacing(4)

        fmt_label = QLabel("出力形式：")
        fmt_layout.addWidget(fmt_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        fmt_layout.addStretch(1)

        fmt_wrap = QWidget()
        fmt_wrap.setContentsMargins(0, 0, 10, 0)
        fmt_wrap_layout = QHBoxLayout(fmt_wrap)
        fmt_wrap_layout.setContentsMargins(0, 0, 0, 0)

        self.format_box = QComboBox()
        self.format_box.addItems(["html", "md", "json", "csv", "pdf"])
        self.format_box.setCurrentText("pdf")
        self.format_box.setFixedWidth(80)
        self.format_box.setFixedHeight(28)
        self.format_box.setToolTip("OCR結果の保存形式を選択します。")
        fmt_wrap_layout.addWidget(self.format_box)

        fmt_layout.addWidget(fmt_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        fmt_widget = QWidget()
        fmt_widget.setLayout(fmt_layout)
        grid.addWidget(fmt_widget, 2, 2, alignment=Qt.AlignVCenter)

        # ▼▼▼ 図の幅(px) ▼▼▼
        width_layout = QHBoxLayout()
        width_layout.setContentsMargins(0, 0, 0, 0)
        width_layout.setSpacing(4)

        width_label = QLabel("図の幅(px)：")
        width_layout.addWidget(width_label, alignment=Qt.AlignLeft | Qt.AlignVCenter)
        width_layout.addStretch(1)

        width_wrap = QWidget()
        width_wrap.setContentsMargins(0, 0, 10, 0)
        width_wrap_layout = QHBoxLayout(width_wrap)
        width_wrap_layout.setContentsMargins(0, 0, 0, 0)

        self.figure_width_input = QLineEdit()
        self.figure_width_input.setValidator(QIntValidator(1, 5000))
        self.figure_width_input.setFixedWidth(100)
        self.figure_width_input.setFixedHeight(28)
        self.figure_width_input.setToolTip(
            "抽出した図の表示幅を指定します。<br>"
            "HTML/Markdown の画像幅指定に使用されます。"
        )
        width_wrap_layout.addWidget(self.figure_width_input)

        width_layout.addWidget(width_wrap, alignment=Qt.AlignRight | Qt.AlignVCenter)

        width_widget = QWidget()
        width_widget.setLayout(width_layout)
        grid.addWidget(width_widget, 2, 3, alignment=Qt.AlignVCenter)

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
        self.log_view.setToolTip("処理の進行状況やエラー内容を表示します。")
        parent_layout.addWidget(self.log_view)

    # --------------------------------------------------------
    # 3-4. UI 有効/無効・進捗表示
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
    # 3-5. Drag & Drop（ファイル追加）
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
        last_dir = self.config.get("Settings", "last_file_dir", fallback="")

        patterns = " ".join(f"*{ext}" for ext in self.SUPPORTED_EXT)
        name_filter = f"対応ファイル ({patterns})"

        files, _ = QFileDialog.getOpenFileNames(
            self,
            "ファイルを選択",
            last_dir,
            name_filter
        )

        if not files:
            return

        folder = str(Path(files[0]).parent)
        self.config["Settings"]["last_file_dir"] = folder
        self.save_config()

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

    # 3-7. 実行処理（スレッド化）
    # --------------------------------------------------------
    def run_yomitoku(self):
        if not self.input_paths:
            self.log("入力が選択されていません。")
            return

        self.total_files = len(self.input_paths)

        exe_path = self.yomitoku_path
        if exe_path is None:
            self.log("YomiToku のパスが設定されていません。")
            return

        fmt = self.format_box.currentText()

        # 出力先
        if getattr(self, "output_dir", None) is None:
            first = self.input_paths[0]
            outdir_path = first.parent if first.is_file() else first
        else:
            outdir_path = self.output_dir

        # DPI
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

        # 書字方向
        ro_map = {
            "自動": "auto",
            "横書き": "left2right",
            "縦書き:上→下": "top2bottom",
            "縦書き:右→左": "right2left"
        }
        reading_order = ro_map[self.direction_box.currentText()]

        # ページ指定
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

        # ▼▼▼ 新しい GUI 項目の取得（ここが重要） ▼▼▼
        figure = self.figure_check.isChecked()
        figure_letter = self.figure_letter_check.isChecked()
        figure_width = self.figure_width_input.text().strip()
        combine = self.combine_check.isChecked()
        ignore_line_break = self.ignore_lb_check.isChecked()
        ignore_meta = self.ignore_meta_check.isChecked()
        encoding = self.encoding_box.currentText()

        # ▼▼▼ 設定ファイルから内部設定を取得 ▼▼▼
        s = self.config["Settings"]
        
        font_path = s.get("font_path", "")
        td_cfg = s.get("td_cfg", "")
        tr_cfg = s.get("tr_cfg", "")
        lp_cfg = s.get("lp_cfg", "")
        tsr_cfg = s.get("tsr_cfg", "")
        figure_dir = s.get("figure_dir", "")

        # ▼▼▼ figure_dir の決定ロジック ▼▼▼
        # ini に設定がある場合はそれを使う
        if figure_dir:
            final_figure_dir = figure_dir
        else:
            # 無い場合は入力ファイルと同じフォルダ
            first_input = self.input_paths[0]
            if first_input.is_file():
                final_figure_dir = str(first_input.parent)
            else:
                final_figure_dir = str(first_input)

        # Worker に渡す値として上書き
        figure_dir = final_figure_dir

        # モード
        lite = self.lite_check.isChecked()

        # ▼▼▼ vis は ini から取得
        vis = False
        if "Settings" in self.config:
            vis = self.config["Settings"].get("vis", "0") == "1"

        # UI を無効化
        self.disable_ui()
        self.reset_run_button()

        # スレッド開始
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
            figure_letter,
            figure_width,
            combine,
            ignore_line_break,
            ignore_meta,
            encoding,
            lite,
            vis,
            self.device,
            font_path,
            td_cfg,
            tr_cfg,
            lp_cfg,
            tsr_cfg,
            figure_dir
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

    # --------------------------------------------------------
    # 4-1. ツールチップの挙動（即表示・無制限）
    # --------------------------------------------------------
    class TooltipStyle(QProxyStyle):
        def styleHint(self, hint, option=None, widget=None, returnData=None):
            if hint == QStyle.SH_ToolTip_WakeUpDelay:
                return 100
            if hint == QStyle.SH_ToolTip_FallAsleepDelay:
                return 0
            return super().styleHint(hint, option, widget, returnData)

    app.setStyle(TooltipStyle())

    # --------------------------------------------------------
    # 4-2. ツールチップの見た目（背景色・文字色・幅など）
    # --------------------------------------------------------
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

    # --------------------------------------------------------
    # 4-3. GUI 起動
    # --------------------------------------------------------
    gui = YomiTokuGUI()
    gui.show()
    sys.exit(app.exec())
