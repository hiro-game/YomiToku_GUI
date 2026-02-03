# ============================================================
# 1. import 群
# ============================================================
import subprocess
import configparser
import sys
from pathlib import Path

from PySide6.QtWidgets import (
    QApplication, QWidget,
    QVBoxLayout, QHBoxLayout, QGridLayout,
    QPushButton, QLabel, QLineEdit,
    QComboBox, QListWidget, QTextEdit,
    QFileDialog
)

from PySide6.QtGui import (
    QIntValidator,QFont,
    QPainter, QColor, QBrush, QPen
)

from PySide6.QtCore import (
    Qt, QThread, QObject, Signal,
    QRect, QSize
)

INI_FILE = "YomiToku_GUI.ini"

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
# 2. Worker クラス（バックエンド処理）
# ============================================================
class YomiTokuWorker(QObject):

    log_signal = Signal(str)
    progress_signal = Signal(int)
    finished = Signal()

    def __init__(self, exe_path, cmd_list):
        """
        exe_path : YomiToku の実行ファイルパス
        cmd_list : create_Option が生成した「引数リスト」の配列
                   例: [
                        ["-o", "out", "-f", "pdf", ... , "input1.pdf"],
                        ["-o", "out", "-f", "pdf", ... , "input2.pdf"],
                      ]
        """
        super().__init__()
        self.exe_path = exe_path
        self.cmd_list = cmd_list

    def run(self):
        total = len(self.cmd_list)
        self.log_signal.emit(f"----- {total} 件の処理を開始 -----")

        for idx, cmd in enumerate(self.cmd_list, start=1):

            # 進捗更新
            progress = int((idx - 1) / total * 100)
            self.progress_signal.emit(progress)

            input_file = cmd[-1]  # create_Option の最後は input_path
            self.log_signal.emit(f"[{idx}/{total}] 処理中: {input_file}")

            # exe_path を先頭に付けた完全コマンド
            full_cmd = [str(self.exe_path)] + cmd

            # ログ出力
            self.log_signal.emit("実行コマンド: " + " ".join(full_cmd))

            # サブプロセス実行
            process = subprocess.Popen(
                full_cmd,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True
            )
            if process.stdout:
                for line in process.stdout:
                    self.log_signal.emit(line.rstrip())

            process.wait()
            self.log_signal.emit(f"完了: {input_file} (終了コード: {process.returncode})")

        # 最終進捗
        self.progress_signal.emit(100)
        self.log_signal.emit("----- 全ての処理が完了しました -----")
        self.finished.emit()

class YomiTokuGUI(QWidget):

    # 対応拡張子
    SUPPORTED_EXT = [".pdf", ".jpg", ".jpeg", ".png", ".bmp", ".tif", ".tiff"]

    # --------------------------------------------------------
    # 3-1. 初期化
    # --------------------------------------------------------
    def __init__(self):
        super().__init__()

        # フォント
        font = QFont()
        font.setPointSize(12)
        self.setFont(font)

        # ini / config
        self.config_path = Path("YomiToku_GUI.ini")
        self.config = configparser.ConfigParser()

        # 内部状態
        self._init_basic_state()

        # UI 構築（★中身は既存コードをそのまま使う）
        self._build_ui()

        # 設定ファイルの確認・初期生成
        self.check_Config()

        # 設定読み込み
        self.load_Fixed()
        self.load_Settings()

        # ウィンドウサイズ固定
        self.adjustSize()
        self.setFixedSize(self.size())

        # 起動ログ
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.log(f"=== App started at {timestamp} ===")

    # --------------------------------------------------------
    # 基本状態
    # --------------------------------------------------------
    def _init_basic_state(self):
        self.setWindowTitle("YomiToku_GUI")
        self.resize(900, 650)
        self.setAcceptDrops(True)

        self.input_paths: list[Path] = []
        self.output_dir: Path | None = None

        # last_dir 系
        self.last_file_dir = ""
        self.last_folder_dir = ""

        # Fixed
        self.yomitoku_path: Path | None = None
        self.device: str = ""

        # Advanced
        self.vis = "0"
        self.td_cfg = ""
        self.tf_cfg = ""   # 新 CLI に合わせて tf_cfg として扱う
        self.lp_cfg = ""
        self.tsr_cfg = ""
        self.figure_dir = ""
        self.font_path = ""

        # Save
        self.save_settings_flag = False
        self.save_log_flag = False

        # 進捗
        self.total_files = 0

    # --------------------------------------------------------
    # ログ表示
    # --------------------------------------------------------
    def log(self, text: str):
        self.log_view.append(text)

    # --------------------------------------------------------
    # ini チェック＆補完
    # --------------------------------------------------------
    def check_Config(self):
        # --------------------------------------------------------
        # 1. ini が無ければ作成（保存は create_Config 内でのみ行われる）
        # --------------------------------------------------------
        if not self.config_path.exists():
            self.create_Config()

        # 常に最新の ini を self.config に読み込む
        self.config.read(self.config_path, encoding="utf-8")

        # --------------------------------------------------------
        # 2. セクション存在チェック（補完するが保存はしない）
        # --------------------------------------------------------
        for sec in ("Fixed", "Settings", "Advanced", "Save"):
            if sec not in self.config:
                self.config[sec] = {}

        # --------------------------------------------------------
        # 3. Fixed の補完（保存は detect 系に任せる）
        # --------------------------------------------------------
        fixed = self.config["Fixed"]

        # yomitoku_path が空なら detect_Path（保存は detect_Path 内）
        if not fixed.get("yomitoku_path", "").strip():
            self.detect_Path()  # save_Fixed により self.config も更新される

        # device が "__AUTO__" または空なら detect_Device（保存は detect_Device 内）
        dev = fixed.get("device", "").strip()
        if dev == "__AUTO__" or not dev:
            self.detect_Device()  # save_Fixed により self.config も更新される

        # --------------------------------------------------------
        # 4. 最後に Fixed / Settings を GUI に反映
        # --------------------------------------------------------
        self.load_Fixed()     # self.config → self
        self.load_Settings()  # self.config → GUI

    # --------------------------------------------------------
    # ini 作成
    # --------------------------------------------------------
    def create_Config(self):
        cfg = configparser.ConfigParser()

        cfg["Fixed"] = {
            "yomitoku_path": "",
            "device": "",
        }

        cfg["Settings"] = {
            "format": "pdf",
            "output_dir": "",
            "lite": "0",
            "ignore_line_break": "0",
            "figure": "0",
            "figure_letter": "0",
            "figure_width": "",
            "encoding": "utf-8",
            "combine": "0",
            "ignore_meta": "0",
            "reading_order": "auto",
            "dpi": "200",
            "pages": "",
            "last_file_dir": "",
            "last_folder_dir": "",
        }

        cfg["Advanced"] = {
            "vis": "0",
            "td_cfg": "",
            "tf_cfg": "",
            "lp_cfg": "",
            "tsr_cfg": "",
            "figure_dir": "",
            "font_path": "",
        }

        cfg["Save"] = {
            "save_settings": "0",
            "save_log": "0",
        }

        with open(self.config_path, "w", encoding="utf-8") as f:
            cfg.write(f)

        self.config = cfg

    # --------------------------------------------------------
    # Fixed 読み込み
    # --------------------------------------------------------
    def load_Fixed(self):
        fixed = self.config["Fixed"]

        yomitoku_path = fixed.get("yomitoku_path", "").strip()
        self.yomitoku_path = Path(yomitoku_path) if yomitoku_path else None

        self.device = fixed.get("device", "").strip()

    # --------------------------------------------------------
    # Settings 読み込み（ini → GUI）
    # --------------------------------------------------------
    def load_Settings(self):
        s = self.config["Settings"]

        self.last_file_dir = s.get("last_file_dir", "")
        self.last_folder_dir = s.get("last_folder_dir", "")

        output_dir = s.get("output_dir", "")
        self.output_dir = Path(output_dir) if output_dir else None

        # format / reading_order / dpi / pages / figure_width / encoding
        if "format" in s:
            value = s["format"]
            index = self.format_box.findData(value)
            if index >= 0:
                self.format_box.setCurrentIndex(index)
                if "reading_order" in s:
                    value = s["reading_order"]
            index = self.direction_box.findData(value)
            if index >= 0:
                self.direction_box.setCurrentIndex(index)
        if "dpi" in s:
            self.dpi_box.setCurrentText(s["dpi"])
        if "pages" in s:
            self.pages_input.setText(s["pages"])
        if "figure_width" in s:
            self.figure_width_input.setText(s["figure_width"])
        if "encoding" in s:
            self.encoding_box.setCurrentText(s["encoding"])

        # チェックボックス
        self.figure_check.setChecked(s.get("figure", "0") == "1")
        self.lite_check.setChecked(s.get("lite", "0") == "1")
        self.figure_letter_check.setChecked(s.get("figure_letter", "0") == "1")
        self.ignore_lb_check.setChecked(s.get("ignore_line_break", "0") == "1")
        self.ignore_meta_check.setChecked(s.get("ignore_meta", "0") == "1")
        self.combine_check.setChecked(s.get("combine", "0") == "1")

    # --------------------------------------------------------
    # Advanced 読み込み（ini → 内部変数 / Advanced UI）
    # --------------------------------------------------------
    def load_Advanced(self):
        parser = configparser.ConfigParser()
        parser.read(self.config_path, encoding="utf-8")

        if "Advanced" not in parser:
            # Advanced セクションが無い場合はデフォルト値
            self.vis = "0"
            self.td_cfg = ""
            self.tf_cfg = ""
            self.lp_cfg = ""
            self.tsr_cfg = ""
            self.figure_dir = ""
            self.font_path = ""
            return

        a = parser["Advanced"]

        self.vis = a.get("vis", "0")
        self.td_cfg = a.get("td_cfg", "")
        self.tf_cfg = a.get("tf_cfg", "")
        self.lp_cfg = a.get("lp_cfg", "")
        self.tsr_cfg = a.get("tsr_cfg", "")
        self.figure_dir = a.get("figure_dir", "")
        self.font_path = a.get("font_path", "")

        # 図保存先 UI がある前提
        if hasattr(self, "figure_dir_input"):
            self.figure_dir_input.setText(self.figure_dir)

    # --------------------------------------------------------
    # Save 読み込み
    # --------------------------------------------------------
    def load_Save(self):
        parser = configparser.ConfigParser()
        parser.read(self.config_path, encoding="utf-8")

        if "Save" not in parser:
            self.save_settings_flag = False
            self.save_log_flag = False
            return

        s = parser["Save"]

        self.save_settings_flag = s.get("save_settings", "0") == "1"
        self.save_log_flag = s.get("save_log", "0") == "1"

    # --------------------------------------------------------
    # Fixed 保存
    # --------------------------------------------------------
    def save_Fixed(self):
        # --- 1) self.config を更新（常に最新の真実にする） ---
        if "Fixed" not in self.config:
            self.config.add_section("Fixed")

        self.config["Fixed"]["yomitoku_path"] = str(self.yomitoku_path) if self.yomitoku_path else ""
        self.config["Fixed"]["device"] = self.device if self.device else ""

        # --- 2) ini を行単位で書き換える ---
        with open(self.config_path, "r", encoding="utf-8") as f:
            lines = f.readlines()

        new_lines = []
        in_fixed = False

        for line in lines:
            stripped = line.strip()

            # Fixed セクション開始
            if stripped == "[Fixed]":
                in_fixed = True
                new_lines.append(line)  # [Fixed] 行そのまま

                # ★ self.config の最新値を書き込む
                new_lines.append(f"yomitoku_path = {self.config['Fixed']['yomitoku_path']}\n")
                new_lines.append(f"device = {self.config['Fixed']['device']}\n")
                continue

            # 次のセクションに入ったら Fixed 終了
            if in_fixed and stripped.startswith("[") and stripped.endswith("]"):
                in_fixed = False

            # Fixed セクション中の古い行はスキップ
            if in_fixed:
                continue

            new_lines.append(line)

        # --- 3) 書き戻し ---
        with open(self.config_path, "w", encoding="utf-8") as f:
            f.writelines(new_lines)

    # --------------------------------------------------------
    # Settings 保存（GUI → ini）
    # --------------------------------------------------------
    def save_Settings(self):
        # Settings セクションの新しい内容（公式ヘルプ順）
        new_settings = []
        new_settings.append("[Settings]\n")
        new_settings.append(f"format = {self.format_box.currentData()}\n")
        new_settings.append(f"output_dir = {self.output_dir or ''}\n")
        new_settings.append(f"lite = {'1' if self.lite_check.isChecked() else '0'}\n")
        new_settings.append(f"ignore_line_break = {'1' if self.ignore_lb_check.isChecked() else '0'}\n")
        new_settings.append(f"figure = {'1' if self.figure_check.isChecked() else '0'}\n")
        new_settings.append(f"figure_letter = {'1' if self.figure_letter_check.isChecked() else '0'}\n")
        new_settings.append(f"figure_width = {self.figure_width_input.text()}\n")
        new_settings.append(f"encoding = {self.encoding_box.currentData()}\n")
        new_settings.append(f"combine = {'1' if self.combine_check.isChecked() else '0'}\n")
        new_settings.append(f"ignore_meta = {'1' if self.ignore_meta_check.isChecked() else '0'}\n")
        new_settings.append(f"reading_order = {self.direction_box.currentData()}\n")
        new_settings.append(f"dpi = {self.dpi_box.currentText()}\n")
        new_settings.append(f"pages = {self.pages_input.text()}\n")
        new_settings.append(f"last_file_dir = {self.last_file_dir}\n")
        new_settings.append(f"last_folder_dir = {self.last_folder_dir}\n")

        # ini 全体を読み込む
        with open(self.config_path, "r", encoding="utf-8") as f:
            lines = f.readlines()

        # 新しい ini を構築
        new_lines = []
        inside_settings = False

        for line in lines:
            stripped = line.strip()

            # Settings セクション開始
            if stripped.lower() == "[settings]":
                inside_settings = True
                new_lines.extend(new_settings)
                continue

            # 次のセクションに入ったら Settings 終了
            if inside_settings and stripped.startswith("[") and stripped.endswith("]"):
                inside_settings = False

            # Settings 内の古い行はスキップ
            if inside_settings:
                continue

            # Settings 以外はそのまま残す
            new_lines.append(line)

            # 書き戻し前に空行を除去
            new_lines = [line for line in new_lines if line.strip() != ""]

            with open(self.config_path, "w", encoding="utf-8") as f:
                f.writelines(new_lines)

        # 書き戻し
        with open(self.config_path, "w", encoding="utf-8") as f:
            f.writelines(new_lines)

    # --------------------------------------------------------
    # YomiToku 実行ファイル検出
    # --------------------------------------------------------
    def detect_Path(self) -> Path | None:

        exe_name = "yomitoku.exe" if sys.platform.startswith("win") else "yomitoku"

        print("detect_Path: start")

        # --------------------------------------------------------
        # 1. pip show -f yomitoku
        # --------------------------------------------------------
        try:
            result = subprocess.check_output(
                ["pip", "show", "-f", "yomitoku"],
                text=True
            )
            location = None
            rel_exe = None

            for line in result.splitlines():
                line = line.strip()

                if line.startswith("Location:"):
                    location = line.split(":", 1)[1].strip()

                # exe_name の末尾一致で誤検出を完全排除
                if line.endswith(exe_name):
                    rel_exe = line
                    break

            if location and rel_exe:
                exe_path = (Path(location) / rel_exe).resolve()
                if exe_path.exists():
                    print(f"detect_Path: found via pip show → {exe_path}")
                    self.yomitoku_path = str(exe_path)
                    self.save_Fixed()  # ★ self.config と ini を同期
                    return exe_path
        except Exception as e:
            print("detect_Path pip show error:", e)

        # --------------------------------------------------------
        # 2. where（Windows）/ which（Linux/Mac）
        # --------------------------------------------------------
        try:
            cmd = ["where", exe_name] if sys.platform.startswith("win") else ["which", exe_name]
            result = subprocess.check_output(cmd, text=True).strip()
            if result:
                p = Path(result)
                if p.exists():
                    print(f"detect_Path: found via where/which → {p}")
                    self.yomitoku_path = str(p)
                    self.save_Fixed()  # ★ 同期
                    return p
        except Exception as e:
            print("detect_Path where/which error:", e)

        # --------------------------------------------------------
        # 3. Scripts/bin フォルダ探索
        # --------------------------------------------------------
        candidates = []

        if sys.platform.startswith("win"):
            scripts_dir = Path(sys.executable).parent / "Scripts"
            candidates.append(scripts_dir / exe_name)
        else:
            candidates.extend([
                Path.home() / ".local/bin" / exe_name,
                Path("/usr/local/bin") / exe_name,
            ])

            if sys.platform == "darwin":
                pyver = f"{sys.version_info.major}.{sys.version_info.minor}"
                candidates.append(Path.home() / f"Library/Python/{pyver}/bin" / exe_name)

        # --------------------------------------------------------
        # 4. site-packages 内の実行ファイル探索
        # --------------------------------------------------------
        try:
            import site
            site_paths = []

            try:
                site_paths.extend(site.getsitepackages())
            except Exception:
                pass

            try:
                site_paths.append(site.getusersitepackages())
            except Exception:
                pass

            for sp in site_paths:
                sp = Path(sp)
                candidates.append(sp / "yomitoku" / exe_name)
                candidates.append(sp / "yomitoku" / "__main__.py")
        except Exception as e:
            print("detect_Path site-packages error:", e)

        # --------------------------------------------------------
        # 5. 候補を順にチェック
        # --------------------------------------------------------
        for c in candidates:
            if c.exists():
                resolved = c.resolve()
                print(f"detect_Path: found via candidates → {resolved}")
                self.yomitoku_path = str(resolved)
                self.save_Fixed()  # ★ 同期
                return resolved

        print("detect_Path: not found")
        self.yomitoku_path = None
        self.save_Fixed()  # ★ None を同期
        return None

    def detect_Device(self) -> str | None:
        # ひとまず空文字を返す実装にしておく
        return ""

    # --------------------------------------------------------
    # ファイル選択（ini には書かない）
    # --------------------------------------------------------
    def select_files(self):
        last_dir = self.last_file_dir or ""

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
        self.last_file_dir = folder

        unique = []
        seen = set()
        for f in files:
            p = Path(f).resolve()
            if p.suffix.lower() in self.SUPPORTED_EXT and p not in seen:
                seen.add(p)
                unique.append(p)

        self.file_list.clear()
        self.input_paths = []
        for p in unique:
            self.file_list.addItem(str(p))
            self.input_paths.append(p)

        self.reset_run_button()

    # --------------------------------------------------------
    # フォルダ選択（ini には書かない）
    # --------------------------------------------------------
    def select_folder(self):
        last_dir = self.last_folder_dir or ""
        folder = QFileDialog.getExistingDirectory(self, "フォルダを選択", last_dir)
        if not folder:
            return

        self.last_folder_dir = folder  # 保存は closeEvent 時

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

    # --------------------------------------------------------
    # 出力先選択（ini には書かない）
    # --------------------------------------------------------
    def select_output(self):
        folder = QFileDialog.getExistingDirectory(self, "出力先フォルダ")
        if folder:
            self.output_dir = Path(folder)

    # --------------------------------------------------------
    # 図保存先選択（Advanced UI）
    # --------------------------------------------------------
    def select_figure_dir(self):
        folder = QFileDialog.getExistingDirectory(self, "図の保存先フォルダ")
        if folder and hasattr(self, "figure_dir_input"):
            self.figure_dir_input.setText(folder)

    # --------------------------------------------------------
    # create_Option（新仕様・条件厳守）
    # --------------------------------------------------------
    def create_Option(self, input_path: Path, base_outdir: Path | None) -> list[str]:
        cmd: list[str] = []

        # -f（format）
        fmt = self.format_box.currentData() or "pdf"
        cmd += ["-f", fmt]

        # -v（vis）
        if self.vis == "1":
            cmd.append("-v")

        # -o（outdir）
        outdir = base_outdir if base_outdir else input_path.parent
        cmd += ["-o", str(outdir)]

        # -l（lite）
        if self.lite_check.isChecked():
            cmd.append("-l")

        # -d（device）
        if self.device:
            cmd += ["-d", self.device]

        # --td_cfg
        if self.td_cfg:
            cmd += ["--td_cfg", self.td_cfg]

        # --tr_cfg
        if self.tf_cfg:
            cmd += ["--tr_cfg", self.tf_cfg]

        # --lp_cfg
        if self.lp_cfg:
            cmd += ["--lp_cfg", self.lp_cfg]

        # --tsr_cfg
        if self.tsr_cfg:
            cmd += ["--tsr_cfg", self.tsr_cfg]

        # --ignore_line_break
        if self.ignore_lb_check.isChecked():
            cmd.append("--ignore_line_break")

        # --figure
        figure_enabled = self.figure_check.isChecked()
        if figure_enabled:
            cmd.append("--figure")

        # --figure_letter（独立オプション）
        if self.figure_letter_check.isChecked():
            cmd.append("--figure_letter")

        # --figure_width
        fig_width = self.figure_width_input.text().strip()
        if figure_enabled and fig_width:
            cmd += ["--figure_width", fig_width]

        # --figure_dir
        if figure_enabled:
            # ini の値をそのまま使う
            fig_dir = getattr(self, "figure_dir", "")

            # ini にも無ければ outdir を使う
            if not fig_dir:
                fig_dir = str(outdir)

            cmd += ["--figure_dir", fig_dir]

        encoding = self.encoding_box.currentData() or "utf-8"
        cmd += ["--encoding", encoding]

        # combine（PDF 入力時のみ）
        is_pdf = input_path.suffix.lower() == ".pdf"
        if is_pdf and self.combine_check.isChecked():
            cmd.append("--combine")

        # --ignore_meta
        if self.ignore_meta_check.isChecked():
            cmd.append("--ignore_meta")

        # --reading_order
        reading_order = self.direction_box.currentData() or "auto"
        cmd += ["--reading_order", reading_order]

        # --font_path（pdf のみ）
        if fmt == "pdf" and self.font_path:
            cmd += ["--font_path", self.font_path]

        # --dpi（pdf のみ）
        dpi_text = self.dpi_box.currentText().strip()
        if is_pdf and dpi_text:
            cmd += ["--dpi", dpi_text]

        # --pages（pdf のみ）
        pages_raw = self.pages_input.text().strip()
        if is_pdf and pages_raw and pages_raw != "0":
            cmd += ["--pages", pages_raw]

        # ★ 最後に入力ファイル（arg1）
        cmd.append(str(input_path))

        return cmd

    # --------------------------------------------------------
    # run_yomitoku
    # --------------------------------------------------------
    def run_yomitoku(self):
        """
        現在の GUI 設定からコマンドライン引数を生成し、
        YomiTokuWorker にまとめて渡して実行する。
        """
        if not self.input_paths:
            self.log("入力が選択されていません。")
            return

        exe_path = getattr(self, "yomitoku_path", None)
        if not exe_path:
            self.log("YomiToku のパスが設定されていません。")
            return

        # Advanced を最新化
        self.load_Advanced()

        self.total_files = len(self.input_paths)

        base_outdir = self.output_dir  # None の場合は create_Option 側で input.parent を使う

        self.disable_ui()
        self.reset_run_button()

        cmd_list: list[list[str]] = []
        for input_path in self.input_paths:
            cmd = self.create_Option(input_path, base_outdir)
            cmd_list.append(cmd)

        self.thread = QThread()
        self.worker = YomiTokuWorker(exe_path, cmd_list)
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

    # --------------------------------------------------------
    # closeEvent
    # --------------------------------------------------------
    def closeEvent(self, event):

        # 終了時に Save セクションを読み込む
        # （起動中に ini を編集して save_settings/save_log が変わっても反映される）
        self.load_Save()

        # Save セクションの値を参照
        save_settings = self.save_settings_flag
        save_log = self.save_log_flag

        # Settings セクションの保存
        if save_settings:
            self.save_Settings()

        # ログウィンドウの保存
        if save_log:
            from datetime import datetime
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            log_name = f"YomiToku_{timestamp}.log"
            with open(log_name, "w", encoding="utf-8") as f:
                f.write(self.log_view.toPlainText())

        event.accept()

    # --------------------------------------------------------
    # Drag & Drop
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

        seen = set(self.input_paths)
        new_items = []

        for url in urls:
            path = Path(url.toLocalFile()).resolve()
            if path.exists() and path.suffix.lower() in self.SUPPORTED_EXT:
                if path not in seen:
                    seen.add(path)
                    new_items.append(path)

        for p in new_items:
            self.input_paths.append(p)
            self.file_list.addItem(str(p))

        self.reset_run_button()

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
        # 1 行目：DPI / ヘッダー等を無視 / 軽量モード / 図を抽出
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
        self.dpi_box.addItems(["200", "400", "600", "1000", "2000"])
        self.dpi_box.setCurrentText("200")
        self.dpi_box.lineEdit().setValidator(QIntValidator(100, 2147483647))
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

        # ▼▼▼ 軽量モード ▼▼▼
        lite_layout = QHBoxLayout()
        lite_layout.setContentsMargins(0, 0, 10, 0)
        lite_layout.setSpacing(4)

        lite_label = QLabel("軽量モード：")
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
            "PDFから読込むページを指定します。\n1P、3P～5Pを読み込む場合\n例：1,3-5　(0か空白で全選択)"
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

        # ★ 表示名と内部値を分離
        items = [
            ("自動", "auto", "内容に応じて自動判定します。\n通常はこちらを選んでください"),
            ("横書き", "left2right", "左から右方向の読み順。\nレシートや保険証など"),
            ("縦書き:上→下", "top2bottom", "上から下方向の読み順。\n段組みの Word 文書など"),
            ("縦書き:右→左", "right2left", "右から左方向の読み順。\n縦書き文書に適しています"),
        ]

        # アイテム追加＋個別ツールチップ設定
        for index, (label, value, tooltip) in enumerate(items):
            self.direction_box.addItem(label, value)
            self.direction_box.setItemData(index, tooltip, Qt.ToolTipRole)

        # デフォルトは「自動」
        self.direction_box.setCurrentIndex(0)

        self.direction_box.setFixedWidth(120)
        self.direction_box.setFixedHeight(28)

        # 全体ツールチップ（補足）
        self.direction_box.setToolTip(
            "画像内テキストの書字方向を指定します。<br>"
            "各項目にカーソルを合わせると詳細が表示されます。"
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
        self.encoding_box = QComboBox()

        items = [
            ("UTF-8", "utf-8"),
            ("UTF-8 BOM", "utf-8-sig"),
            ("SHIFT-JIS", "shift-jis"),
            ("EUC-JP", "euc-jp"),
            ("Windows-31J", "cp932"),
        ]

        for label, value in items:
            self.encoding_box.addItem(label, value)
        self.encoding_box.setItemData(self.encoding_box.findData("utf-8-sig"), "UTF-8 BOM", Qt.ToolTipRole)
        self.encoding_box.setItemData(self.encoding_box.findData("cp932"), "Windows-31J", Qt.ToolTipRole)
        self.encoding_box.setCurrentIndex(0)

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
            "複数ページのPDFを読み取った場合、通常は複数のファイルに出力されますが、このオプションは単一のファイルにまとめて出力します。"
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

        # ★ 表示名と内部値を分離
        items = [
            ("HTML", "html"),
            ("Markdown", "md"),
            ("JSON", "json"),
            ("CSV", "csv"),
            ("PDF", "pdf"),
        ]

        for label, value in items:
            self.format_box.addItem(label, value)

        # デフォルトは PDF
        index = self.format_box.findData("pdf")
        if index >= 0:
            self.format_box.setCurrentIndex(index)

        self.format_box.setItemData(self.format_box.findData("md"), "MarkDown", Qt.ToolTipRole)
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

    def _remove_file_item(self, item):
        row = self.file_list.row(item)
        self.file_list.takeItem(row)

        if 0 <= row < len(self.input_paths):
            del self.input_paths[row]

    def refresh_file_list(self):
        self.file_list.clear()
        for p in self.input_paths:
            self.file_list.addItem(str(p))

# ============================================================
# 4. Main
# ============================================================
if __name__ == "__main__":
    app = QApplication(sys.argv)

    gui = YomiTokuGUI()
    gui.show()

    sys.exit(app.exec())