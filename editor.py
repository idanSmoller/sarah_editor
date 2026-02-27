import sys
import subprocess
import tempfile
import shutil
from pathlib import Path
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QSlider, QFileDialog,
                             QLabel, QStyle, QSizePolicy, QFrame,
                             QMessageBox, QProgressDialog, QDialog)
from PyQt6.QtMultimedia import QMediaPlayer, QAudioOutput
from PyQt6.QtMultimediaWidgets import QVideoWidget
from PyQt6.QtCore import Qt, QUrl, QSize
from PyQt6.QtGui import QFont


STYLESHEET = """
    QMainWindow, QWidget {
        background-color: #1e1e1e;
        color: #f0f0f0;
    }
    QVideoWidget {
        background-color: #000000;
    }
    QSlider::groove:horizontal {
        height: 6px;
        background: #444;
        border-radius: 3px;
    }
    QSlider::handle:horizontal {
        background: #ffffff;
        width: 14px;
        height: 14px;
        margin: -4px 0;
        border-radius: 7px;
    }
    QSlider::sub-page:horizontal {
        background: #0078d4;
        border-radius: 3px;
    }
    QPushButton {
        background-color: #2d2d2d;
        color: #f0f0f0;
        border: 1px solid #555;
        border-radius: 6px;
        padding: 6px 16px;
        font-size: 13px;
    }
    QPushButton:hover {
        background-color: #3a3a3a;
        border-color: #888;
    }
    QPushButton:pressed {
        background-color: #0078d4;
        border-color: #0078d4;
    }
    QPushButton#play_button {
        background-color: #0078d4;
        border-color: #0078d4;
        border-radius: 20px;
        padding: 0px;
    }
    QPushButton#play_button:hover {
        background-color: #1a8fe3;
    }
    QPushButton#start_button {
        background-color: #107c10;
        border-color: #107c10;
        font-weight: bold;
    }
    QPushButton#start_button:hover { background-color: #1a9e1a; }
    QPushButton#stop_and_start_button {
        background-color: #ca5010;
        border-color: #ca5010;
        font-weight: bold;
    }
    QPushButton#stop_and_start_button:hover { background-color: #e05a12; }
    QPushButton#stop_button {
        background-color: #c42b1c;
        border-color: #c42b1c;
        font-weight: bold;
    }
    QPushButton#stop_button:hover { background-color: #d93025; }
    QLabel#time_label {
        color: #aaaaaa;
        font-size: 12px;
    }
    QFrame#divider {
        color: #444;
    }
"""


class VideoEditor(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Video Editor")
        self.setGeometry(100, 100, 1080, 700)
        self.setMinimumSize(640, 480)

        # Data storage
        self.segments = []
        self.video_path = None
        self.frame_ms = int(1000 / 30)  # default; updated once video metadata loads

        # OneDrive state
        self.is_onedrive = False
        self.onedrive_client = None
        self.onedrive_parent_id = None
        self.stream_url = None

        self.setStyleSheet(STYLESHEET)

        # Setup UI
        self.setup_ui()

        # Load video file
        self.load_video()

    def setup_ui(self):
        # Central widget
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(0)

        # Video widget — takes all available vertical space
        self.video_widget = QVideoWidget()
        self.video_widget.setObjectName("video_widget")
        self.video_widget.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        layout.addWidget(self.video_widget, stretch=1)

        # Media player setup
        self.media_player = QMediaPlayer()
        self.audio_output = QAudioOutput()
        self.media_player.setAudioOutput(self.audio_output)
        self.media_player.setVideoOutput(self.video_widget)

        # ── Bottom control bar ──────────────────────────────────────────────
        control_bar = QWidget()
        control_bar.setFixedHeight(110)
        control_bar.setStyleSheet("background-color: #252525;")
        bar_layout = QVBoxLayout(control_bar)
        bar_layout.setContentsMargins(16, 10, 16, 10)
        bar_layout.setSpacing(8)

        # Time slider
        self.time_slider = QSlider(Qt.Orientation.Horizontal)
        self.time_slider.sliderMoved.connect(self.set_position)
        bar_layout.addWidget(self.time_slider)

        # Row: time label + buttons
        row = QHBoxLayout()
        row.setSpacing(10)

        # Time label (left-aligned)
        self.time_label = QLabel("00:00 / 00:00")
        self.time_label.setObjectName("time_label")
        self.time_label.setFixedWidth(130)
        row.addWidget(self.time_label)

        row.addStretch()

        # Play/Pause button (round, icon-only)
        self.play_button = QPushButton()
        self.play_button.setObjectName("play_button")
        self.play_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
        self.play_button.setIconSize(QSize(22, 22))
        self.play_button.setFixedSize(42, 42)
        self.play_button.clicked.connect(self.play_pause)
        row.addWidget(self.play_button)

        row.addStretch()

        # Segment buttons (right-aligned)
        self.start_button = QPushButton("▶  Start")
        self.start_button.setObjectName("start_button")
        self.start_button.setFixedHeight(36)
        self.start_button.clicked.connect(self.on_start)
        row.addWidget(self.start_button)

        self.stop_and_start_button = QPushButton("⏭  Stop & Start")
        self.stop_and_start_button.setObjectName("stop_and_start_button")
        self.stop_and_start_button.setFixedHeight(36)
        self.stop_and_start_button.clicked.connect(self.on_stop_and_start)
        row.addWidget(self.stop_and_start_button)

        self.stop_button = QPushButton("■  Stop")
        self.stop_button.setObjectName("stop_button")
        self.stop_button.setFixedHeight(36)
        self.stop_button.clicked.connect(self.on_stop)
        row.addWidget(self.stop_button)

        bar_layout.addLayout(row)
        layout.addWidget(control_bar)

        # Initial state: not recording → show only Start
        self.is_recording = False
        self._update_button_state()

        # Connect media player signals
        self.media_player.positionChanged.connect(self.position_changed)
        self.media_player.durationChanged.connect(self.duration_changed)
        self.media_player.mediaStatusChanged.connect(self.on_media_status_changed)
    
    def load_video(self):
        msg = QMessageBox(self)
        msg.setWindowTitle("Open Video")
        msg.setText("Where would you like to open a video from?")
        local_btn = msg.addButton("📂  Local File", QMessageBox.ButtonRole.AcceptRole)
        onedrive_btn = msg.addButton("☁️  OneDrive", QMessageBox.ButtonRole.AcceptRole)
        msg.addButton("Cancel", QMessageBox.ButtonRole.RejectRole)
        msg.exec()

        clicked = msg.clickedButton()
        if clicked == local_btn:
            self._load_local()
        elif clicked == onedrive_btn:
            self._load_onedrive()

    def _load_local(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Video File", "", "Video Files (*.mp4 *.avi *.mov *.mkv)"
        )
        if file_path:
            self.video_path = file_path
            self.is_onedrive = False
            self.media_player.setSource(QUrl.fromLocalFile(file_path))

    def _load_onedrive(self):
        try:
            from config import AZURE_CLIENT_ID
            if AZURE_CLIENT_ID == "YOUR_CLIENT_ID_HERE":
                QMessageBox.warning(
                    self, "Not Configured",
                    "Please set AZURE_CLIENT_ID in config.py first.\n"
                    "Open config.py for step-by-step instructions."
                )
                return

            from onedrive import OneDriveClient
            from onedrive_browser import OneDriveBrowser

            client = OneDriveClient(AZURE_CLIENT_ID)
            client.authenticate()

            browser = OneDriveBrowser(client, self)
            if browser.exec() == QDialog.DialogCode.Accepted and browser.selected_item:
                item = browser.selected_item
                self.onedrive_client = client
                self.onedrive_parent_id = item["parent_id"]
                self.video_path = item["name"]
                self.is_onedrive = True
                self.stream_url = client.get_download_url(item["id"])
                self.media_player.setSource(QUrl(self.stream_url))
        except Exception as e:
            QMessageBox.critical(self, "OneDrive Error", str(e))
    
    def play_pause(self):
        if self.media_player.playbackState() == QMediaPlayer.PlaybackState.PlayingState:
            self.media_player.pause()
            self.play_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
        else:
            self.media_player.play()
            self.play_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPause))
    
    def set_position(self, position):
        self.media_player.setPosition(position)
    
    def position_changed(self, position):
        self.time_slider.setValue(position)
        self.update_time_label(position, self.media_player.duration())
    
    def duration_changed(self, duration):
        self.time_slider.setRange(0, duration)
    
    def update_time_label(self, position, duration):
        position_time = self.format_time(position)
        duration_time = self.format_time(duration)
        self.time_label.setText(f"{position_time} / {duration_time}")
    
    def format_time(self, milliseconds):
        seconds = milliseconds // 1000
        minutes = seconds // 60
        seconds = seconds % 60
        return f"{minutes:02d}:{seconds:02d}"
    
    def _update_button_state(self):
        """Show Start XOR (Stop and Start + Stop) depending on recording state."""
        self.start_button.setVisible(not self.is_recording)
        self.stop_and_start_button.setVisible(self.is_recording)
        self.stop_button.setVisible(self.is_recording)

    def on_start(self):
        """Begin a new segment at the current position."""
        current_time = self.media_player.position()
        self.segments.append({"start": current_time, "stop": None})
        self.is_recording = True
        self._update_button_state()
        print(f"Start point added at {self.format_time(current_time)}")

    def on_stop(self):
        """Close the current segment — switches back to Start button."""
        current_time = self.media_player.position()
        self._close_current_segment(current_time)
        self.is_recording = False
        self._update_button_state()

    def on_stop_and_start(self):
        """Close the current segment and immediately open a new one."""
        current_time = self.media_player.position()
        self._close_current_segment(current_time)
        self.segments.append({"start": current_time, "stop": None})
        print(f"New start point added at {self.format_time(current_time)}")
        # Stays in recording state — buttons don't change

    def _close_current_segment(self, current_time):
        """Attach a stop time to the most recent open segment."""
        for segment in reversed(self.segments):
            if segment["stop"] is None:
                segment["stop"] = current_time
                print(f"Stop point added at {self.format_time(current_time)}")
                return
        print("Warning: no open segment to close")
    
    def on_media_status_changed(self, status):
        """Once the video is loaded, read its frame rate from metadata."""
        if status == QMediaPlayer.MediaStatus.LoadedMedia:
            from PyQt6.QtMultimedia import QMediaMetaData
            fps = self.media_player.metaData().value(QMediaMetaData.Key.VideoFrameRate)
            if fps and fps > 0:
                self.frame_ms = int(1000 / fps)
                print(f"Detected frame rate: {fps:.2f} fps ({self.frame_ms} ms/frame)")
            else:
                print("Frame rate not found in metadata, using 30 fps default")

    def keyPressEvent(self, event):
        """Left/right arrow keys step one frame backward/forward."""
        if event.key() == Qt.Key.Key_Right:
            self.media_player.pause()
            self.play_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
            self.media_player.setPosition(min(self.media_player.position() + self.frame_ms,
                                              self.media_player.duration()))
        elif event.key() == Qt.Key.Key_Left:
            self.media_player.pause()
            self.play_button.setIcon(self.style().standardIcon(QStyle.StandardPixmap.SP_MediaPlay))
            self.media_player.setPosition(max(self.media_player.position() - self.frame_ms, 0))
        else:
            super().keyPressEvent(event)

    def closeEvent(self, event):
        """On close, slice video into clips; upload to OneDrive if in cloud mode."""
        # Auto-close any open segment at the end of the video
        total_duration = self.media_player.duration()
        for segment in self.segments:
            if segment["stop"] is None:
                segment["stop"] = total_duration
                print(f"Auto-closed open segment at {self.format_time(total_duration)}")

        complete_segments = [s for s in self.segments if s["stop"] is not None]

        if complete_segments and self.video_path:
            video_stem = Path(self.video_path).stem

            if self.is_onedrive:
                output_dir = Path(tempfile.mkdtemp())
                video_input = self.stream_url
            else:
                video_path_obj = Path(self.video_path)
                output_dir = video_path_obj.parent / f"{video_stem}_clips"
                output_dir.mkdir(exist_ok=True)
                video_input = str(video_path_obj)

            clip_files = []
            for i, segment in enumerate(complete_segments):
                start_sec = segment["start"] / 1000.0
                stop_sec = segment["stop"] / 1000.0
                clip_duration = stop_sec - start_sec
                clip_name = f"{video_stem}_cut_{i + 1:02d}.mp4"
                output_file = output_dir / clip_name

                cmd = [
                    "ffmpeg", "-y",
                    "-ss", str(start_sec),
                    "-i", video_input,
                    "-t", str(clip_duration),
                    "-c", "copy",
                    str(output_file)
                ]
                print(f"Exporting clip {i + 1}: {self.format_time(segment['start'])} → {self.format_time(segment['stop'])}")
                subprocess.run(cmd, check=True)
                clip_files.append((clip_name, output_file))

            if self.is_onedrive:
                self._upload_clips_to_onedrive(clip_files, video_stem, output_dir)
            else:
                print(f"\nAll clips saved to: {output_dir}")

        event.accept()

    def _upload_clips_to_onedrive(self, clip_files, video_stem, tmp_dir):
        """Upload exported clips to a new OneDrive folder, then clean up temp files."""
        try:
            folder_id = self.onedrive_client.create_folder(
                self.onedrive_parent_id, f"{video_stem}_clips"
            )
            progress = QProgressDialog(
                "Uploading clips to OneDrive...", "Cancel", 0, len(clip_files), self
            )
            progress.setWindowTitle("Uploading")
            progress.setWindowModality(Qt.WindowModality.WindowModal)
            progress.show()

            for i, (clip_name, clip_path) in enumerate(clip_files):
                if progress.wasCanceled():
                    break
                progress.setLabelText(f"Uploading {clip_name}  ({i + 1} / {len(clip_files)})")
                progress.setValue(i)
                QApplication.processEvents()
                self.onedrive_client.upload_file(folder_id, clip_name, clip_path)

            progress.setValue(len(clip_files))
            print(f"All clips uploaded to OneDrive: {video_stem}_clips/")
        except Exception as e:
            QMessageBox.critical(self, "Upload Error", f"Failed to upload clips:\n{e}")
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)


def main():
    app = QApplication(sys.argv)
    editor = VideoEditor()
    editor.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
