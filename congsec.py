# -*- coding: utf-8 -*-
# congsec_gui.py
import sys
import json
import os
import re
import csv
from datetime import datetime
from PyQt5.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QTextEdit, QListWidget, QListWidgetItem, QLabel,
    QSpinBox, QTabWidget, QFileDialog, QMessageBox, QProgressBar,
    QGroupBox, QFrame, QDialog, QDialogButtonBox, QCheckBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QTextCharFormat, QColor, QSyntaxHighlighter
from PyQt5.QtWidgets import QPlainTextEdit

# -------------------- 高亮显示类 --------------------
class ResultHighlighter(QSyntaxHighlighter):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.highlighting_rules = []

        keyword_format = QTextCharFormat()
        keyword_format.setBackground(QColor(255, 255, 0))
        self.highlighting_rules.append((re.compile(r"关键字列表: (.+)"), keyword_format))

        match_format = QTextCharFormat()
        match_format.setBackground(QColor(255, 255, 0))
        self.highlighting_rules.append((re.compile(r"匹配到 \d+ 个关键字列表"), match_format))

        label_format = QTextCharFormat()
        label_format.setBackground(QColor(255, 255, 200))
        self.highlighting_rules.append((re.compile(r"附近行内容:|附近文字:|文件:|-{50}|排除文本:|向下行内容:|向上行内容:"), label_format))

        exclude_format = QTextCharFormat()
        exclude_format.setBackground(QColor(255, 200, 200))
        self.highlighting_rules.append((re.compile(r"已排除（.*）"), exclude_format))

    def highlightBlock(self, text):
        for pattern, fmt in self.highlighting_rules:
            for match in pattern.finditer(text):
                start, end = match.span()
                self.setFormat(start, end - start, fmt)

# -------------------- 全屏结果显示窗口 --------------------
class FullscreenResultWindow(QDialog):
    def __init__(self, parent=None, text="", title="全屏结果"):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setWindowState(Qt.WindowMaximized)
        layout = QVBoxLayout(self)
        self.text_edit = QPlainTextEdit()
        self.text_edit.setPlainText(text)
        self.text_edit.setReadOnly(True)
        self.highlighter = ResultHighlighter(self.text_edit.document())
        layout.addWidget(self.text_edit)
        close_btn = QPushButton("关闭")
        close_btn.clicked.connect(self.accept)
        layout.addWidget(close_btn)

# -------------------- 工作线程类 --------------------
class WorkerThread(QThread):
    progress_signal = pyqtSignal(int, int, str)
    result_signal = pyqtSignal(str, list)
    finished_signal = pyqtSignal()
    error_signal = pyqtSignal(str)

    def __init__(self, config, files):
        super().__init__()
        self.config = config
        self.files = files
        self.is_running = True
        self.chunk_size = 1024 * 1024  # 1MB chunks for large files

    def run(self):
        try:
            total_files = len(self.files)
            all_results = []
            result_texts = []
            for i, file_path in enumerate(self.files):
                if not self.is_running:
                    break
                self.progress_signal.emit(i + 1, total_files, os.path.basename(file_path))
                try:
                    content = self.read_file_optimized(file_path)
                    if content is None:
                        continue  # Skip binary files

                    result_text, file_results = self.process_text(content, self.config, f"文件: {os.path.basename(file_path)}")
                    all_results.extend(file_results)
                    result_texts.append(result_text)
                except Exception as e:
                    self.error_signal.emit(f"读取文件 {file_path} 时出错: {e}")

            full_result_text = "\n".join(result_texts)
            full_result_text = f"处理完成！共处理 {total_files} 个文件\n\n" + full_result_text
            self.result_signal.emit(full_result_text, all_results)
        except Exception as e:
            self.error_signal.emit(f"处理过程中出错: {e}")
        finally:
            self.finished_signal.emit()

    def stop(self):
        self.is_running = False

    def read_file_optimized(self, file_path):
        try:
            with open(file_path, 'rb') as f:
                chunk = f.read(1024)
                if b'\x00' in chunk:
                    return None  # Binary file, skip

            file_size = os.path.getsize(file_path)
            if file_size > 10 * 1024 * 1024:  # For files larger than 10MB
                content = []
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    while True:
                        chunk = f.read(self.chunk_size)
                        if not chunk:
                            break
                        content.append(chunk)
                        if not self.is_running:
                            return None
                return ''.join(content)
            else:
                with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                    return f.read()
        except UnicodeDecodeError:
            return None  # Skip files that can't be decoded as text
        except Exception as e:
            raise e

    def process_text(self, text, config, source="unknown"):
        keywords = config["keywords"]
        nearby_lines = config["nearby_lines"]
        nearby_chars = config["nearby_chars"]
        results = []
        result_lines = []
        total_hits = 0
        lines = text.splitlines()

        for kw in keywords:
            words = kw.get("words", [])
            exclude = kw.get("exclude", [])
            kw_lines = kw.get("nearby_lines", nearby_lines)
            kw_chars = kw.get("nearby_chars", nearby_chars)
            down_lines = kw.get("down_lines", 0)
            up_lines = kw.get("up_lines", 0)
            exclude_nearby = kw.get("exclude_nearby", True)
            multi_line_exclude = kw.get("multi_line_exclude", False)

            for line_no, line in enumerate(lines, 1):
                if not words:
                    continue

                start_line = max(0, line_no - 1 - kw_lines)
                end_line = min(len(lines), line_no + kw_lines)
                nearby_lines_text = "\n".join(lines[start_line:end_line])

                down_text = ""
                if down_lines != 0:
                    if down_lines > 0:
                        down_start = line_no
                        down_end = min(len(lines), line_no + down_lines)
                    else:
                        down_start = max(0, line_no + down_lines - 1)
                        down_end = line_no
                    down_text = "\n".join(lines[down_start:down_end])

                up_text = ""
                if up_lines != 0:
                    if up_lines > 0:
                        up_start = max(0, line_no - 1 - up_lines)
                        up_end = line_no - 1
                    else:
                        up_start = line_no - 1
                        up_end = min(len(lines), line_no - 1 - up_lines)
                    up_text = "\n".join(lines[up_start:up_end])

                nearby_chars_text = ""
                if kw_chars > 0:
                    pattern = re.compile(r'(' + '|'.join(re.escape(word) for word in words[0:1]) + ')')
                    matches = list(pattern.finditer(line))
                    if matches:
                        parts = []
                        for match in matches:
                            start = match.start()
                            end = match.end()
                            pre_start = max(0, start - kw_chars)
                            pre_end = start
                            post_start = end
                            post_end = min(len(line), end + kw_chars)
                            pre_part = line[pre_start:pre_end]
                            post_part = line[post_start:post_end]
                            parts.append(f"{pre_part}[{match.group()}]{post_part}")
                        seen = set()
                        unique_parts = []
                        for part in parts:
                            if part not in seen:
                                seen.add(part)
                                unique_parts.append(part)
                        nearby_chars_text = "\n".join(unique_parts)

                match_success = False
                if multi_line_exclude:
                    if any(w in line for w in words[0:1]):
                        other_keywords = words[1:]
                        if other_keywords:
                            combined_content = (
                                nearby_lines_text + "\n" +
                                nearby_chars_text + "\n" +
                                down_text + "\n" +
                                up_text
                            )
                            all_found = all(word in combined_content for word in other_keywords)
                            if all_found:
                                match_success = True
                            else:
                                continue
                        else:
                            match_success = True
                else:
                    all_found = all(word in nearby_chars_text for word in words)
                    if all_found:
                        match_success = True

                if not match_success:
                    continue

                combined_text = line
                if exclude_nearby:
                    combined_text += nearby_lines_text + nearby_chars_text + down_text + up_text

                excluded = any(e and e in combined_text for e in exclude)
                if excluded:
                    result_lines.append(f"已排除（包含排除文本）: {' + '.join(words)}（位于第 {line_no} 行）")
                    result_lines.append("-" * 50)
                    continue

                total_hits += 1
                result_lines.append(f"关键字列表: {' + '.join(words)}（位于第 {line_no} 行）")
                result_lines.append("附近行内容:")
                result_lines.append(nearby_lines_text)
                if kw_chars > 0:
                    result_lines.append("附近文字:")
                    result_lines.append(nearby_chars_text)
                if down_lines != 0:
                    direction = "向下" if down_lines > 0 else "向上"
                    result_lines.append(f"{direction}行内容:")
                    result_lines.append(down_text)
                if up_lines != 0:
                    direction = "向上" if up_lines > 0 else "向下"
                    result_lines.append(f"{direction}行内容:")
                    result_lines.append(up_text)
                result_lines.append("-" * 50)

                results.append({
                    "keywords": " + ".join(words),
                    "line_number": line_no,
                    "nearby_lines": nearby_lines_text,
                    "nearby_chars": nearby_chars_text,
                    "down_lines": down_text,
                    "up_lines": up_text,
                    "source": source,
                    "exclude_text": "; ".join(exclude)
                })

        header = f"匹配到 {total_hits} 个关键字列表"
        result_lines.insert(0, header)
        result_text = "\n".join(result_lines)
        return result_text, results

# -------------------- 主窗口类 --------------------
class CongsecGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config = self.load_config()
        self.worker_thread = None
        self.current_results = []
        self.result_buffer = []
        self.buffer_timer = QTimer()
        self.buffer_timer.timeout.connect(self.flush_buffer)
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("文本批量处理工具 by CongSec~")
        self.setGeometry(100, 100, 1200, 800)
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QHBoxLayout(central_widget)

        # 左侧配置面板
        config_panel = QFrame()
        config_panel.setFrameStyle(QFrame.Box)
        config_panel.setMaximumWidth(300)
        config_layout = QVBoxLayout(config_panel)

        keyword_group = QGroupBox("关键字列表")
        keyword_layout = QVBoxLayout(keyword_group)
        self.keyword_list = QListWidget()
        self.update_keyword_list()
        keyword_layout.addWidget(self.keyword_list)

        add_keyword_btn = QPushButton("添加关键字")
        add_keyword_btn.clicked.connect(self.add_keyword_dialog)
        keyword_layout.addWidget(add_keyword_btn)

        edit_keyword_btn = QPushButton("编辑选中关键字")
        edit_keyword_btn.clicked.connect(self.edit_keyword_dialog)
        keyword_layout.addWidget(edit_keyword_btn)

        delete_keyword_btn = QPushButton("删除选中关键字")
        delete_keyword_btn.clicked.connect(self.delete_keyword)
        keyword_layout.addWidget(delete_keyword_btn)
        config_layout.addWidget(keyword_group)

        config_group = QGroupBox("默认配置")
        config_group_layout = QVBoxLayout(config_group)

        lines_layout = QHBoxLayout()
        lines_layout.addWidget(QLabel("默认附近行数:"))
        self.lines_spin = QSpinBox()
        self.lines_spin.setRange(0, 100)
        self.lines_spin.setValue(self.config["nearby_lines"])
        self.lines_spin.valueChanged.connect(self.update_default_config)
        lines_layout.addWidget(self.lines_spin)
        config_group_layout.addLayout(lines_layout)

        chars_layout = QHBoxLayout()
        chars_layout.addWidget(QLabel("默认附近字符数:"))
        self.chars_spin = QSpinBox()
        self.chars_spin.setRange(0, 1000)
        self.chars_spin.setValue(self.config["nearby_chars"])
        chars_layout.addWidget(self.chars_spin)
        config_group_layout.addLayout(chars_layout)

        down_layout = QHBoxLayout()
        down_layout.addWidget(QLabel("默认向下行数:"))
        self.down_spin = QSpinBox()
        self.down_spin.setRange(-100, 100)
        self.down_spin.setValue(self.config.get("down_lines", 0))
        self.down_spin.valueChanged.connect(self.update_default_config)
        down_layout.addWidget(self.down_spin)
        config_group_layout.addLayout(down_layout)

        up_layout = QHBoxLayout()
        up_layout.addWidget(QLabel("默认向上行数:"))
        self.up_spin = QSpinBox()
        self.up_spin.setRange(-100, 100)
        self.up_spin.setValue(self.config.get("up_lines", 0))
        self.up_spin.valueChanged.connect(self.update_default_config)
        up_layout.addWidget(self.up_spin)
        config_group_layout.addLayout(up_layout)

        # ★★★ 新增：后台自动导出开关
        self.auto_export_cb = QCheckBox("后台自动导出CSV")
        self.auto_export_cb.setChecked(self.config.get("auto_export", True))
        self.auto_export_cb.toggled.connect(self.toggle_auto_export)
        config_group_layout.addWidget(self.auto_export_cb)

        config_layout.addWidget(config_group)
        config_layout.addStretch()

        # 右侧主区域
        right_panel = QVBoxLayout()
        self.tab_widget = QTabWidget()

        # 批量处理标签页
        batch_tab = QWidget()
        batch_layout = QVBoxLayout(batch_tab)

        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout(file_group)

        select_files_btn = QPushButton("选择文件")
        select_files_btn.clicked.connect(self.select_files)
        file_layout.addWidget(select_files_btn)

        select_folder_btn = QPushButton("选择文件夹(递归)")
        select_folder_btn.clicked.connect(self.select_folder_recursive)
        file_layout.addWidget(select_folder_btn)

        self.selected_files_label = QLabel("未选择文件")
        file_layout.addWidget(self.selected_files_label)
        batch_layout.addWidget(file_group)

        self.progress_bar = QProgressBar()
        self.progress_bar.setVisible(False)
        batch_layout.addWidget(self.progress_bar)
        self.progress_label = QLabel("")
        self.progress_label.setVisible(False)
        batch_layout.addWidget(self.progress_label)

        self.show_excluded_cb_batch = QCheckBox("显示已排除提示")
        self.show_excluded_cb_batch.setChecked(False)
        batch_layout.addWidget(self.show_excluded_cb_batch)

        process_btn = QPushButton("开始批量处理")
        process_btn.clicked.connect(self.start_batch_processing)
        batch_layout.addWidget(process_btn)

        self.stop_btn = QPushButton("停止处理")
        self.stop_btn.clicked.connect(self.stop_processing)
        self.stop_btn.setVisible(False)
        batch_layout.addWidget(self.stop_btn)

        self.export_csv_btn = QPushButton("导出结果到CSV")
        self.export_csv_btn.clicked.connect(self.export_to_csv)
        self.export_csv_btn.setVisible(False)
        batch_layout.addWidget(self.export_csv_btn)

        result_group = QGroupBox("匹配结果")
        result_layout = QVBoxLayout(result_group)
        self.result_text = QPlainTextEdit()
        self.result_text.setReadOnly(True)
        self.highlighter = ResultHighlighter(self.result_text.document())
        result_layout.addWidget(self.result_text)

        fullscreen_batch_btn = QPushButton("全屏查看")
        fullscreen_batch_btn.clicked.connect(self.show_batch_fullscreen)
        result_layout.addWidget(fullscreen_batch_btn)
        batch_layout.addWidget(result_group)
        self.tab_widget.addTab(batch_tab, "批量处理")

        # 实时处理标签页
        realtime_tab = QWidget()
        realtime_layout = QVBoxLayout(realtime_tab)

        input_group = QGroupBox("输入文本")
        input_layout = QVBoxLayout(input_group)
        self.input_text = QPlainTextEdit()
        self.input_text.setPlaceholderText("请输入要匹配的文本...")
        input_layout.addWidget(self.input_text)

        self.show_excluded_cb_realtime = QCheckBox("显示已排除提示")
        self.show_excluded_cb_realtime.setChecked(False)
        input_layout.addWidget(self.show_excluded_cb_realtime)

        process_realtime_btn = QPushButton("单个匹配")
        process_realtime_btn.clicked.connect(self.process_realtime)
        input_layout.addWidget(process_realtime_btn)
        realtime_layout.addWidget(input_group)

        result_group_realtime = QGroupBox("匹配结果")
        result_layout_realtime = QVBoxLayout(result_group_realtime)
        self.result_text_realtime = QPlainTextEdit()
        self.result_text_realtime.setReadOnly(True)
        self.highlighter_realtime = ResultHighlighter(self.result_text_realtime.document())
        result_layout_realtime.addWidget(self.result_text_realtime)

        fullscreen_realtime_btn = QPushButton("全屏查看")
        fullscreen_realtime_btn.clicked.connect(self.show_realtime_fullscreen)
        result_layout_realtime.addWidget(fullscreen_realtime_btn)

        export_realtime_btn = QPushButton("导出结果到CSV")
        export_realtime_btn.clicked.connect(self.export_realtime_to_csv)
        result_layout_realtime.addWidget(export_realtime_btn)
        realtime_layout.addWidget(result_group_realtime)
        self.tab_widget.addTab(realtime_tab, "单个匹配")

        right_panel.addWidget(self.tab_widget)
        main_layout.addWidget(config_panel)
        main_layout.addLayout(right_panel)

    def load_config(self):
        config_path = "config.json"
        default_config = {
            "keywords": [],
            "nearby_lines": 2,
            "nearby_chars": 20,
            "down_lines": 0,
            "up_lines": 0,
            "auto_export": True  # ★★★ 新增
        }
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    for idx, kw in enumerate(config.get("keywords", [])):
                        if isinstance(kw, str):
                            config["keywords"][idx] = {
                                "words": [kw],
                                "exclude": [],
                                "enabled": True,
                                "down_lines": 0,
                                "up_lines": 0,
                                "exclude_nearby": True,
                                "multi_line_exclude": False
                            }
                        elif isinstance(kw, dict):
                            if "word" in kw:
                                kw["words"] = [kw.pop("word")]
                            kw.setdefault("exclude", [])
                            kw.setdefault("enabled", True)
                            kw.setdefault("down_lines", 0)
                            kw.setdefault("up_lines", 0)
                            kw.setdefault("exclude_nearby", True)
                            kw.setdefault("multi_line_exclude", False)
                    for key in default_config:
                        config.setdefault(key, default_config[key])
                    return config
            except Exception as e:
                QMessageBox.warning(self, "配置错误", f"读取配置文件出错: {e}，使用默认配置")
                return default_config.copy()
        else:
            with open(config_path, 'w', encoding='utf-8') as f:
                json.dump(default_config, f, ensure_ascii=False, indent=4)
            return default_config.copy()

    def save_config(self):
        with open("config.json", 'w', encoding='utf-8') as f:
            json.dump(self.config, f, ensure_ascii=False, indent=4)

    def update_keyword_list(self):
        self.keyword_list.clear()
        for kw in self.config["keywords"]:
            words = kw.get("words", [])
            exclude = kw.get("exclude", [])
            lines = kw.get("nearby_lines", self.config["nearby_lines"])
            chars = kw.get("nearby_chars", self.config["nearby_chars"])
            down = kw.get("down_lines", self.config["down_lines"])
            up = kw.get("up_lines", self.config["up_lines"])
            text = f"关键字: {'+'.join(words)} (行:{lines} 字符:{chars} 下:{down} 上:{up})"
            if exclude:
                text += f" | 排除: {'/'.join(exclude)}"
            if kw.get("multi_line_exclude", False):
                text += " | 多行过滤: 是"
            item = QListWidgetItem(text)
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Checked if kw.get("enabled", True) else Qt.Unchecked)
            self.keyword_list.addItem(item)

    def update_default_config(self):
        self.config["nearby_lines"] = self.lines_spin.value()
        self.config["nearby_chars"] = self.chars_spin.value()
        self.config["down_lines"] = self.down_spin.value()
        self.config["up_lines"] = self.up_spin.value()
        self.save_config()
        self.update_keyword_list()

    # ★★★ 新增：切换自动导出
    def toggle_auto_export(self, checked):
        self.config["auto_export"] = checked
        self.save_config()

    def add_keyword_dialog(self):
        dialog = QDialog(self)
        dialog.setWindowTitle("添加关键字")
        dialog.setModal(True)
        layout = QVBoxLayout(dialog)

        keyword_edit = QPlainTextEdit()
        keyword_edit.setMaximumHeight(80)
        exclude_edit = QPlainTextEdit()
        exclude_edit.setMaximumHeight(80)

        lines_spin = QSpinBox()
        lines_spin.setRange(0, 100)
        lines_spin.setValue(self.config["nearby_lines"])

        chars_spin = QSpinBox()
        chars_spin.setRange(0, 1000)
        chars_spin.setValue(self.config["nearby_chars"])

        down_spin = QSpinBox()
        down_spin.setRange(-100, 100)
        down_spin.setValue(self.config["down_lines"])

        up_spin = QSpinBox()
        up_spin.setRange(-100, 100)
        up_spin.setValue(self.config["up_lines"])

        exclude_nearby_cb = QCheckBox("排除文本参与附近检查")
        exclude_nearby_cb.setChecked(True)

        multi_line_cb = QCheckBox("多行关键字参与附近匹配过滤")
        multi_line_cb.setToolTip("勾选后，除第一行外的其他关键字如果在附近内容中出现，将排除该结果")

        layout.addWidget(QLabel("关键字（每行/逗号分隔，需全部匹配）:"))
        layout.addWidget(keyword_edit)
        layout.addWidget(QLabel("附近行数:"))
        layout.addWidget(lines_spin)
        layout.addWidget(QLabel("附近字符数:"))
        layout.addWidget(chars_spin)
        layout.addWidget(QLabel("向下行数:"))
        layout.addWidget(down_spin)
        layout.addWidget(QLabel("向上行数:"))
        layout.addWidget(up_spin)
        layout.addWidget(exclude_nearby_cb)
        layout.addWidget(multi_line_cb)
        layout.addWidget(QLabel("排除文本（每行/逗号分隔，命中其一即排除）:"))
        layout.addWidget(exclude_edit)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        if dialog.exec_() == QDialog.Accepted:
            words = [w.strip() for w in re.split(r'[,\n]', keyword_edit.toPlainText()) if w.strip()]
            exclude = [e.strip() for e in re.split(r'[,\n]', exclude_edit.toPlainText()) if e.strip()]
            if words:
                new_keyword = {
                    "words": words,
                    "exclude": exclude,
                    "nearby_lines": lines_spin.value(),
                    "nearby_chars": chars_spin.value(),
                    "down_lines": down_spin.value(),
                    "up_lines": up_spin.value(),
                    "enabled": True,
                    "exclude_nearby": exclude_nearby_cb.isChecked(),
                    "multi_line_exclude": multi_line_cb.isChecked()
                }
                self.config["keywords"].append(new_keyword)
                self.save_config()
                self.update_keyword_list()

    def edit_keyword_dialog(self):
        current_row = self.keyword_list.currentRow()
        if current_row < 0 or current_row >= len(self.config["keywords"]):
            QMessageBox.warning(self, "警告", "请先选择一个关键字进行编辑")
            return
        kw = self.config["keywords"][current_row]

        dialog = QDialog(self)
        dialog.setWindowTitle("编辑关键字")
        dialog.setModal(True)
        layout = QVBoxLayout(dialog)

        keyword_edit = QPlainTextEdit()
        keyword_edit.setPlainText("\n".join(kw.get("words", [])))
        keyword_edit.setMaximumHeight(80)

        exclude_edit = QPlainTextEdit()
        exclude_edit.setPlainText("\n".join(kw.get("exclude", [])))
        exclude_edit.setMaximumHeight(80)

        lines_spin = QSpinBox()
        lines_spin.setRange(0, 100)
        lines_spin.setValue(kw.get("nearby_lines", self.config["nearby_lines"]))

        chars_spin = QSpinBox()
        chars_spin.setRange(0, 1000)
        chars_spin.setValue(kw.get("nearby_chars", self.config["nearby_chars"]))

        down_spin = QSpinBox()
        down_spin.setRange(-100, 100)
        down_spin.setValue(kw.get("down_lines", self.config["down_lines"]))

        up_spin = QSpinBox()
        up_spin.setRange(-100, 100)
        up_spin.setValue(kw.get("up_lines", self.config["up_lines"]))

        exclude_nearby_cb = QCheckBox("排除文本参与附近检查")
        exclude_nearby_cb.setChecked(kw.get("exclude_nearby", True))

        multi_line_cb = QCheckBox("多行关键字参与附近匹配过滤")
        multi_line_cb.setChecked(kw.get("multi_line_exclude", False))
        multi_line_cb.setToolTip("勾选后，除第一行外的其他关键字如果在附近内容中出现，将排除该结果")

        layout.addWidget(QLabel("关键字（每行/逗号分隔，需全部匹配）:"))
        layout.addWidget(keyword_edit)
        layout.addWidget(QLabel("附近行数:"))
        layout.addWidget(lines_spin)
        layout.addWidget(QLabel("附近字符数:"))
        layout.addWidget(chars_spin)
        layout.addWidget(QLabel("向下行数:"))
        layout.addWidget(down_spin)
        layout.addWidget(QLabel("向上行数:"))
        layout.addWidget(up_spin)
        layout.addWidget(exclude_nearby_cb)
        layout.addWidget(multi_line_cb)
        layout.addWidget(QLabel("排除文本（每行/逗号分隔，命中其一即排除）:"))
        layout.addWidget(exclude_edit)

        button_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        button_box.accepted.connect(dialog.accept)
        button_box.rejected.connect(dialog.reject)
        layout.addWidget(button_box)

        if dialog.exec_() == QDialog.Accepted:
            words = [w.strip() for w in re.split(r'[,\n]', keyword_edit.toPlainText()) if w.strip()]
            exclude = [e.strip() for e in re.split(r'[,\n]', exclude_edit.toPlainText()) if e.strip()]
            if words:
                self.config["keywords"][current_row] = {
                    "words": words,
                    "exclude": exclude,
                    "nearby_lines": lines_spin.value(),
                    "nearby_chars": chars_spin.value(),
                    "down_lines": down_spin.value(),
                    "up_lines": up_spin.value(),
                    "enabled": kw.get("enabled", True),
                    "exclude_nearby": exclude_nearby_cb.isChecked(),
                    "multi_line_exclude": multi_line_cb.isChecked()
                }
                self.save_config()
                self.update_keyword_list()

    def delete_keyword(self):
        current_row = self.keyword_list.currentRow()
        if current_row >= 0 and current_row < len(self.config["keywords"]):
            reply = QMessageBox.question(self, "确认删除", "确定要删除这个关键字吗？", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                self.config["keywords"].pop(current_row)
                self.save_config()
                self.update_keyword_list()

    def _enabled_keywords(self):
        enabled = []
        for row, kw in enumerate(self.config["keywords"]):
            item = self.keyword_list.item(row)
            if item.checkState() == Qt.Checked:
                enabled.append(kw)
        return enabled

    def select_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "选择文件", "",
            "All Files (*);;Text Files (*.txt);;Log Files (*.log);;CSV Files (*.csv)"
        )
        if files:
            self.selected_files = files
            self.selected_files_label.setText(f"已选择 {len(files)} 个文件")

    def select_folder_recursive(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            self.selected_files = []
            for root, _, files in os.walk(folder):
                for file in files:
                    file_path = os.path.join(root, file)
                    self.selected_files.append(file_path)
            self.selected_files_label.setText(f"已选择 {len(self.selected_files)} 个文件 (递归搜索)")

    def start_batch_processing(self):
        if not hasattr(self, 'selected_files') or not self.selected_files:
            QMessageBox.warning(self, "警告", "请先选择要处理的文件")
            return
        self.progress_bar.setVisible(True)
        self.progress_label.setVisible(True)
        self.stop_btn.setVisible(True)
        self.export_csv_btn.setVisible(False)
        self.result_text.clear()

        enabled_config = {
            "keywords": self._enabled_keywords(),
            "nearby_lines": self.config["nearby_lines"],
            "nearby_chars": self.config["nearby_chars"],
            "down_lines": self.config["down_lines"],
            "up_lines": self.config["up_lines"]
        }

        self.worker_thread = WorkerThread(enabled_config, self.selected_files)
        self.worker_thread.progress_signal.connect(self.update_progress)
        self.worker_thread.result_signal.connect(self.show_batch_results)
        self.worker_thread.error_signal.connect(self.show_error)
        self.worker_thread.finished_signal.connect(self.processing_finished)
        self.worker_thread.start()

    def stop_processing(self):
        if self.worker_thread and self.worker_thread.isRunning():
            self.worker_thread.stop()
            self.worker_thread.wait()

    def update_progress(self, current, total, filename):
        self.progress_bar.setMaximum(total)
        self.progress_bar.setValue(current)
        self.progress_label.setText(f"正在处理: {filename} ({current}/{total})")

    def show_batch_results(self, result_text, results):
        if not self.show_excluded_cb_batch.isChecked():
            lines = [line for line in result_text.splitlines() if not line.startswith("已排除（")]
            cleaned, prev_separator = [], False
            for line in lines:
                stripped = line.strip()
                if stripped == "-" * 50:
                    if not prev_separator:
                        cleaned.append(line)
                        prev_separator = True
                    continue
                cleaned.append(line)
                prev_separator = False
            result_text = "\n".join(cleaned)

        self.current_results = results
        self.result_text.setPlainText(result_text)
        self.export_csv_btn.setVisible(len(results) > 0)

        # 后台静默导出
        if self.config.get("auto_export", True) and results:
            self.auto_export_results(results, "batch")

    def show_error(self, error_msg):
        QMessageBox.critical(self, "错误", error_msg)

    def processing_finished(self):
        self.progress_bar.setVisible(False)
        self.progress_label.setVisible(False)
        self.stop_btn.setVisible(False)

    def flush_buffer(self):
        if not self.result_buffer:
            self.buffer_timer.stop()
            # 实时匹配结束，自动导出
            if self.config.get("auto_export", True) and self.current_results:
                self.auto_export_results(self.current_results, "realtime")
            return

        chunk = self.result_buffer[:100]
        self.result_buffer = self.result_buffer[100:]
        cursor = self.result_text_realtime.textCursor()
        cursor.movePosition(cursor.End)
        cursor.insertText("\n".join(chunk) + "\n")
        if not self.result_buffer:
            self.buffer_timer.stop()
            # 实时匹配结束，自动导出
            if self.config.get("auto_export", True) and self.current_results:
                self.auto_export_results(self.current_results, "realtime")

    def process_realtime(self):
        text = self.input_text.toPlainText().strip()
        if not text:
            QMessageBox.warning(self, "警告", "请输入要匹配的文本")
            return

        self.result_buffer = []
        self.buffer_timer.start(100)

        enabled_config = {
            "keywords": self._enabled_keywords(),
            "nearby_lines": self.config["nearby_lines"],
            "nearby_chars": self.config["nearby_chars"],
            "down_lines": self.config["down_lines"],
            "up_lines": self.config["up_lines"]
        }

        worker = WorkerThread(enabled_config, [])
        result_text, results = worker.process_text(text, enabled_config, "实时输入")

        if not self.show_excluded_cb_realtime.isChecked():
            lines = [line for line in result_text.splitlines() if not line.startswith("已排除（")]
            cleaned, prev_separator = [], False
            for line in lines:
                stripped = line.strip()
                if stripped == "-" * 50:
                    if not prev_separator:
                        cleaned.append(line)
                        prev_separator = True
                    continue
                cleaned.append(line)
                prev_separator = False
            result_text = "\n".join(cleaned)

        self.current_results = results
        self.result_buffer = result_text.splitlines()
        self.result_text_realtime.clear()
        self.buffer_timer.start(100)

    # 新增：后台自动导出逻辑
    def auto_export_results(self, results, prefix):
        """
        将结果静默导出到 data 目录，文件名：prefix_YYYYMMDD_HHMMSS.csv
        prefix = batch / realtime
        """
        try:
            os.makedirs("data", exist_ok=True)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = os.path.join("data", f"{prefix}_{timestamp}.csv")
            with open(filename, "w", newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=[
                    'keywords', 'line_number', 'nearby_lines', 'nearby_chars',
                    'down_lines', 'up_lines', 'source', 'exclude_text'
                ])
                writer.writeheader()
                for row in results:
                    writer.writerow(row)
        except Exception as e:
            pass

    def export_to_csv(self):
        if not self.current_results:
            QMessageBox.warning(self, "警告", "没有结果可导出")
            return

        filename, _ = QFileDialog.getSaveFileName(
            self, "导出CSV",
            f"batch_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            "CSV Files (*.csv)"
        )

        if filename:
            try:
                export_thread = ExportThread(self.current_results, filename)
                export_thread.start()
                export_thread.wait()
                QMessageBox.information(self, "成功", f"结果已导出到: {filename}")
            except Exception as e:
                QMessageBox.critical(self, "错误", f"导出失败: {e}")

    def export_realtime_to_csv(self):
        self.export_to_csv()

    def show_batch_fullscreen(self):
        text = self.result_text.toPlainText()
        if not text.strip():
            QMessageBox.warning(self, "警告", "没有结果可全屏查看")
            return
        dialog = FullscreenResultWindow(self, text, "批量处理结果 - 全屏")
        dialog.exec_()

    def show_realtime_fullscreen(self):
        text = self.result_text_realtime.toPlainText()
        if not text.strip():
            QMessageBox.warning(self, "警告", "没有结果可全屏查看")
            return
        dialog = FullscreenResultWindow(self, text, "实时处理结果 - 全屏")
        dialog.exec_()


class ExportThread(QThread):
    def __init__(self, results, filename):
        super().__init__()
        self.results = results
        self.filename = filename

    def run(self):
        try:
            with open(self.filename, 'w', newline='', encoding='utf-8-sig') as csvfile:
                fieldnames = ['keywords', 'line_number', 'nearby_lines', 'nearby_chars',
                            'down_lines', 'up_lines', 'source', 'exclude_text']
                writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                writer.writeheader()
                for result in self.results:
                    writer.writerow(result)
        except Exception as e:
            raise e


def main():
    os.environ["QT_DEVICE_PIXEL_RATIO"] = "0"
    os.environ["QT_AUTO_SCREEN_SCALE_FACTOR"] = "1"
    os.environ["QT_SCREEN_SCALE_FACTORS"] = "1"
    os.environ["QT_SCALE_FACTOR"] = "1"
    app = QApplication(sys.argv)
    os.makedirs("data", exist_ok=True)
    window = CongsecGUI()
    window.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
