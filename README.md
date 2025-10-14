# Bulk Media Compressor — Images & Videos

A PyQt5-based desktop tool to **compress images and videos in bulk**, generate an Excel report, and provide real-time logs with drag-and-drop folder selection.

---

## Features

- Compress images (`.jpg`, `.jpeg`, `.png`, `.bmp`, `.tiff`, `.gif`) with adjustable quality.
- Compress videos (`.mp4`, `.avi`, `.mov`, `.mkv`, `.wmv`, `.flv`) with adjustable scale.
- Drag-and-drop folder selection or browse via dialog.
- Threaded processing with progress bar to keep the UI responsive.
- Real-time log updates for each file.
- Cancel button to gracefully stop ongoing compression.
- Generates an Excel report with:
  - Original size
  - Compressed size
  - Size difference (MB and %)
- Automatically opens the Excel report after completion.
- Light-themed GUI with modern buttons and input fields.

---

## Installation

1. Clone the repository:

```bash
git clone <repository-url>
cd bulk-media-compressor
