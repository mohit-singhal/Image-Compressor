import os
import sys
import cv2
import shutil
import openpyxl
from PyQt5 import QtWidgets, QtGui
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from openpyxl.styles import PatternFill, Border, Side, Alignment
import subprocess  # for opening Excel file automatically

# ------------------------ Helper / Excel Utilities ------------------------ #

IMAGE_EXTS = (".png", ".jpg", ".jpeg", ".tiff", ".bmp", ".gif")
VIDEO_EXTS = (".mp4", ".avi", ".mov", ".mkv", ".wmv", ".flv")


def set_column_widths(sheet):
    col_widths = {"B": 6, "C": 33, "D": 12.29, "E": 16.43, "F": 11, "G": 12.29}
    for col, w in col_widths.items():
        sheet.column_dimensions[col].width = w


def merge_and_style_cells(sheet, start_row, start_col, end_row, end_col, text, color):
    rng = f"{sheet.cell(row=start_row, column=start_col).coordinate}:{sheet.cell(row=end_row, column=end_col).coordinate}"
    sheet.merge_cells(rng)
    cell = sheet.cell(row=start_row, column=start_col)
    cell.value = text
    cell.alignment = Alignment(horizontal="center", vertical="center")

    fill = PatternFill(start_color=color, end_color=color, fill_type="solid")
    border = Border(
        left=Side(style="medium", color="000000"),
        right=Side(style="medium", color="000000"),
        top=Side(style="medium", color="000000"),
        bottom=Side(style="medium", color="000000"),
    )

    for row in sheet.iter_rows(min_row=start_row, max_row=end_row, min_col=start_col, max_col=end_col):
        for c in row:
            c.fill = fill
            c.border = border


def get_file_stats(original_folder, compressed_folder, filename):
    orig_path = os.path.join(original_folder, filename)
    comp_path = os.path.join(compressed_folder, filename)

    try:
        orig = os.path.getsize(orig_path) / (1024 * 1024)
    except OSError:
        orig = 0.0
    try:
        comp = os.path.getsize(comp_path) / (1024 * 1024)
    except OSError:
        comp = 0.0

    diff = orig - comp
    pct = (diff * 100 / orig) if orig else 0.0
    return [filename, round(orig, 2), round(comp, 2), round(diff, 2), round(pct, 2)]


# ---------------------------- Worker Thread ---------------------------- #

class CompressorWorker(QThread):
    progress = pyqtSignal(int)
    log = pyqtSignal(str)
    finished = pyqtSignal(str)
    error = pyqtSignal(str)

    def __init__(self, input_folder, compression_percent, parent=None):
        super().__init__(parent)
        self.input_folder = input_folder
        self.compression_percent = compression_percent
        self._is_interrupted = False

    def interrupt(self):
        self._is_interrupted = True

    def _compress_image(self, src_path, dest_path):
        ext = os.path.splitext(src_path)[1].lower()
        img = cv2.imread(src_path)
        if img is None:
            raise RuntimeError("cv2.imread returned None")

        if ext in (".jpg", ".jpeg"):
            cv2.imwrite(dest_path, img, [cv2.IMWRITE_JPEG_QUALITY, self.compression_percent])
        elif ext == ".png":
            png_level = max(0, min(9, int((100 - self.compression_percent) / 10)))
            cv2.imwrite(dest_path, img, [cv2.IMWRITE_PNG_COMPRESSION, png_level])
        else:
            cv2.imwrite(dest_path, img)

    def _compress_video(self, src_path, dest_path):
        cap = cv2.VideoCapture(src_path)
        if not cap.isOpened():
            raise RuntimeError("Unable to open video")

        fps = cap.get(cv2.CAP_PROP_FPS) or 25.0
        frame_count = int(cap.get(cv2.CAP_PROP_FRAME_COUNT) or 0)
        width = int(cap.get(cv2.CAP_PROP_FRAME_WIDTH) or 0)
        height = int(cap.get(cv2.CAP_PROP_FRAME_HEIGHT) or 0)

        scale = max(0.1, self.compression_percent / 100.0)
        new_w = max(1, int(width * scale))
        new_h = max(1, int(height * scale))

        dest_root, _ = os.path.splitext(dest_path)
        dest_mp4 = dest_root + ".mp4"
        fourcc = cv2.VideoWriter_fourcc(*"mp4v")
        writer = cv2.VideoWriter(dest_mp4, fourcc, fps, (new_w, new_h))
        if not writer.isOpened():
            cap.release()
            raise RuntimeError("Unable to open VideoWriter")

        for f_idx in range(frame_count):
            if self._is_interrupted:
                writer.release()
                cap.release()
                raise RuntimeError("Interrupted by user")
            ret, frame = cap.read()
            if not ret:
                break
            if scale != 1.0:
                frame = cv2.resize(frame, (new_w, new_h), interpolation=cv2.INTER_AREA)
            writer.write(frame)

        writer.release()
        cap.release()

        if os.path.exists(dest_mp4):
            if os.path.abspath(dest_mp4) != os.path.abspath(dest_path):
                try:
                    shutil.move(dest_mp4, dest_path)
                    final_path = dest_path
                except Exception:
                    final_path = dest_mp4
            else:
                final_path = dest_mp4
            return final_path
        else:
            raise RuntimeError("Video encoding produced no output file")

    def run(self):
        try:
            src = self.input_folder
            parent_path, name = os.path.split(src)
            dest = os.path.join(parent_path, f"{name}_Compressed")
            os.makedirs(dest, exist_ok=True)

            wb = openpyxl.Workbook()
            sheet = wb.active
            set_column_widths(sheet)
            merge_and_style_cells(sheet, 3, 2, 4, 7, f"Source Path : {src}", "76933C")
            merge_and_style_cells(sheet, 6, 2, 7, 7, f"Destination Path : {dest}", "F79646")

            headers = ["#", "File Name", "Original Size", "Compressed Size", "Difference", "Difference %"]
            header_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
            border = Border(
                left=Side(style="medium", color="000000"),
                right=Side(style="medium", color="000000"),
                top=Side(style="medium", color="000000"),
                bottom=Side(style="medium", color="000000"),
            )

            for col_index, h in enumerate(headers, start=2):
                cell = sheet.cell(row=10, column=col_index)
                cell.value = h
                cell.fill = header_fill
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")

            files = os.listdir(src)
            total_files = len(files) or 1
            processed = 0
            media_counter = 0
            orig_total = 0.0
            comp_total = 0.0
            excel_row = 11

            self.log.emit(f"Found {len(files)} files. Starting compression at {self.compression_percent}%.")

            for idx, fname in enumerate(files, start=1):
                if self._is_interrupted:
                    self.log.emit("Compression interrupted by user.")
                    break

                processed += 1
                src_path = os.path.join(src, fname)
                dest_path = os.path.join(dest, fname)

                if os.path.isdir(src_path):
                    self.log.emit(f"Skipped directory: {fname}")
                    self.progress.emit(int((processed / total_files) * 100))
                    continue

                try:
                    lower = fname.lower()
                    if lower.endswith(IMAGE_EXTS):
                        try:
                            self._compress_image(src_path, dest_path)
                            self.log.emit(f"Compressed image: {fname}")
                        except Exception as e:
                            shutil.copyfile(src_path, dest_path)
                            self.log.emit(f"Image copy fallback: {fname}. Error: {e}")
                    elif lower.endswith(VIDEO_EXTS):
                        try:
                            final_video_path = self._compress_video(src_path, dest_path)
                            dest_path = final_video_path
                            self.log.emit(f"Compressed video: {fname}")
                        except Exception as e:
                            try:
                                shutil.copyfile(src_path, dest_path)
                                self.log.emit(f"Video copy fallback: {fname}. Error: {e}")
                            except Exception as ex:
                                self.log.emit(f"Failed to copy video: {fname}. Error: {ex}")
                    else:
                        try:
                            shutil.copyfile(src_path, dest_path)
                            self.log.emit(f"Copied non-media: {fname}")
                        except Exception as e:
                            self.log.emit(f"Failed to copy {fname}: {e}")

                    stats = get_file_stats(src, dest, os.path.basename(dest_path))
                    if stats:
                        orig_total += stats[1]
                        comp_total += stats[2]

                        row_data = [
                            media_counter + 1,
                            stats[0],
                            f"{stats[1]} MB",
                            f"{stats[2]} MB",
                            f"{stats[3]} MB",
                            f"{stats[4]}%",
                        ]
                        for cidx, val in enumerate(row_data, start=2):
                            cell = sheet.cell(row=excel_row, column=cidx)
                            cell.value = val
                            cell.border = border
                            if cidx != 3:
                                cell.alignment = Alignment(horizontal="center", vertical="center")
                        excel_row += 1
                        media_counter += 1

                except Exception as e:
                    self.log.emit(f"Unexpected error with file {fname}: {e}")

                self.progress.emit(int((processed / total_files) * 100))

            if media_counter > 0:
                diff_mb = round(orig_total - comp_total, 2)
                diff_pct = round((diff_mb * 100) / orig_total, 2) if orig_total else 0
                footer = ["", "Total", f"{orig_total} MB", f"{comp_total} MB", f"{diff_mb} MB", f"{diff_pct}%"]
                footer_fill = PatternFill(start_color="92CDDC", end_color="92CDDC", fill_type="solid")
                for cidx, val in enumerate(footer, start=2):
                    cell = sheet.cell(row=excel_row + 1, column=cidx)
                    cell.value = val
                    cell.fill = footer_fill
                    cell.border = border
                    cell.alignment = Alignment(horizontal="center", vertical="center")

            excel_path = os.path.join(dest, "compression_details.xlsx")
            wb.save(excel_path)

            self.progress.emit(100)
            self.log.emit(f"Compression finished. Excel saved at: {excel_path}")
            self.finished.emit(excel_path)

        except Exception as e:
            self.error.emit(f"Worker failed with error: {e}")


# ------------------------------ GUI Class ------------------------------ #

class FileCompressorUI(QtWidgets.QWidget):
    """Main UI with drag-and-drop folder selection and light theme."""

    def __init__(self):
        super().__init__()
        self.worker = None
        self._setup_ui()
        self._apply_light_theme()
        self.setAcceptDrops(True)

    def _setup_ui(self):
        self.setWindowTitle("Bulk Media Compress Tool — Images & Videos")
        self.setWindowIcon(QtGui.QIcon("icon.png"))
        self.resize(750, 520)

        self.folder_label = QtWidgets.QLabel("Select Folder:")
        self.folder_edit = QtWidgets.QLineEdit()
        self.folder_edit.setPlaceholderText("Drag a folder here or click Browse")
        self.folder_btn = QtWidgets.QPushButton("Browse")
        self.compress_label = QtWidgets.QLabel("Compression % (10-90):")
        self.compress_edit = QtWidgets.QLineEdit()
        self.compress_edit.setPlaceholderText("e.g., 60")
        self.start_btn = QtWidgets.QPushButton("Start")
        self.cancel_btn = QtWidgets.QPushButton("Cancel")
        self.cancel_btn.setEnabled(False)

        self.progress_bar = QtWidgets.QProgressBar()
        self.log_box = QtWidgets.QTextEdit()
        self.log_box.setReadOnly(True)

        layout = QtWidgets.QGridLayout()
        layout.addWidget(self.folder_label, 0, 0)
        layout.addWidget(self.folder_edit, 0, 1, 1, 4)
        layout.addWidget(self.folder_btn, 0, 5)
        layout.addWidget(self.compress_label, 1, 0)
        layout.addWidget(self.compress_edit, 1, 1)
        layout.addWidget(self.start_btn, 1, 3)
        layout.addWidget(self.cancel_btn, 1, 4)
        layout.addWidget(QtWidgets.QLabel("Progress:"), 2, 0)
        layout.addWidget(self.progress_bar, 2, 1, 1, 5)
        layout.addWidget(QtWidgets.QLabel("Logs:"), 3, 0)
        layout.addWidget(self.log_box, 4, 0, 1, 6)
        self.setLayout(layout)

        self.folder_btn.clicked.connect(self.select_folder)
        self.start_btn.clicked.connect(self.start_worker)
        self.cancel_btn.clicked.connect(self.cancel_worker)

    def _apply_light_theme(self):
        self.setStyleSheet("""
            QWidget { background-color: #f9f9f9; font-family: Arial; font-size: 11pt; }
            QLineEdit, QTextEdit { background-color: #ffffff; border: 1px solid #cccccc; border-radius: 4px; padding: 4px; }
            QPushButton { background-color: #e0e0e0; border: 1px solid #b0b0b0; padding: 6px 12px; border-radius: 4px; }
            QPushButton:hover { background-color: #d0d0d0; }
            QProgressBar { border: 1px solid #b0b0b0; border-radius: 5px; text-align: center; }
            QProgressBar::chunk { background-color: #76c7c0; }
        """)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                if os.path.isdir(url.toLocalFile()):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        for url in urls:
            folder = url.toLocalFile()
            if os.path.isdir(folder):
                self.folder_edit.setText(folder)
                self.append_log(f"Folder selected via drag-and-drop: {folder}")
                break

    def select_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Select Folder")
        if folder:
            self.folder_edit.setText(folder)

    def append_log(self, message):
        self.log_box.append(message)
        self.log_box.ensureCursorVisible()

    def start_worker(self):
        folder = self.folder_edit.text().strip()
        pct_text = self.compress_edit.text().strip()

        if not folder or not os.path.exists(folder):
            QtWidgets.QMessageBox.warning(self, "Warning", "Please select a valid folder.")
            return
        if not pct_text.isdigit():
            QtWidgets.QMessageBox.warning(self, "Warning", "Enter a valid compression percentage (10–90).")
            return

        pct = int(pct_text)
        if not (10 <= pct <= 90):
            QtWidgets.QMessageBox.warning(self, "Warning", "Compression % must be between 10 and 90.")
            return

        self.start_btn.setEnabled(False)
        self.cancel_btn.setEnabled(True)
        self.folder_btn.setEnabled(False)
        self.append_log(f"Starting compression: folder={folder}, quality={pct}%")

        self.worker = CompressorWorker(folder, pct)
        self.worker.progress.connect(self.on_progress)
        self.worker.log.connect(self.append_log)
        self.worker.finished.connect(self.on_finished)
        self.worker.error.connect(self.on_error)
        self.worker.start()

    def cancel_worker(self):
        if self.worker and self.worker.isRunning():
            self.worker.interrupt()
            self.append_log("Cancellation requested. Waiting for the worker to stop...")
            self.cancel_btn.setEnabled(False)

    def on_progress(self, value):
        self.progress_bar.setValue(value)

    def on_finished(self, excel_path):
        self.append_log(f"Worker finished. Excel: {excel_path}")
        QtWidgets.QMessageBox.information(self, "Done", f"Compression completed.\nExcel saved at:\n{excel_path}")
        # Open Excel automatically
        try:
            if sys.platform.startswith('darwin'):
                subprocess.call(['open', excel_path])
            elif os.name == 'nt':
                os.startfile(excel_path)
            elif os.name == 'posix':
                subprocess.call(['xdg-open', excel_path])
        except Exception:
            pass
        self._reset_ui()

    def on_error(self, msg):
        self.append_log(f"ERROR: {msg}")
        QtWidgets.QMessageBox.critical(self, "Error", msg)
        self._reset_ui()

    def _reset_ui(self):
        self.start_btn.setEnabled(True)
        self.cancel_btn.setEnabled(False)
        self.folder_btn.setEnabled(True)
        self.progress_bar.setValue(0)


# ------------------------------- Entry Point ------------------------------- #

if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    window = FileCompressorUI()
    window.show()
    sys.exit(app.exec_())
