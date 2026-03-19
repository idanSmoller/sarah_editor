import sys
import subprocess
import shutil
import traceback
import platform
import json
import csv
from pathlib import Path
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout,
                             QHBoxLayout, QPushButton, QSlider, QFileDialog,
                             QLabel, QStyle, QSizePolicy, QMessageBox,
                             QProgressDialog)
from PyQt5.QtMultimedia import QMediaPlayer, QMediaContent
from PyQt5.QtMultimediaWidgets import QVideoWidget
from PyQt5.QtCore import Qt, QUrl, QSize, QObject, QThread, pyqtSignal, QTimer
from PyQt5.QtGui import QPainter, QColor, QBrush


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
    QPushButton#undo_button {
        background-color: #555555;
        border-color: #555555;
    }
    QPushButton#undo_button:hover { background-color: #666666; }
    QLabel#time_label {
        color: #aaaaaa;
        font-size: 12px;
    }
    QFrame#divider {
        color: #444;
    }
"""


class SegmentBar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setFixedHeight(12)
        # Use transparent background for the widget itself, so only the drawn rectangles show
        # or a very subtle background for the track
        self.setAutoFillBackground(False)
        self.segments = []
        self.duration = 0
        self.colors = [
            QColor(255, 0, 0, 100),    # Red
            QColor(0, 255, 0, 100),    # Green
            QColor(0, 0, 255, 100),    # Blue
            QColor(255, 255, 0, 100),  # Yellow
            QColor(255, 0, 255, 100),  # Magenta
            QColor(0, 255, 255, 100),  # Cyan
            QColor(255, 128, 0, 100),  # Orange
            QColor(128, 0, 255, 100),  # Purple
        ]

    def set_data(self, segments, duration):
        self.segments = segments
        self.duration = duration
        self.update()  # Trigger repaint

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing, False)
        
        # Draw background track
        painter.fillRect(self.rect(), QColor(60, 60, 60))

        if not self.duration or self.duration <= 0:
            return

        width = self.width()

        for i, segment in enumerate(self.segments):
            start = segment.get('start', 0)
            stop = segment.get('stop')
            
            # If segment is currently recording (stop is None), skip visualizing for now
            if stop is None:
                continue
            
            # Clamp values to duration
            curr_start = max(0, min(start, self.duration))
            curr_stop = max(0, min(stop, self.duration))
            
            if curr_start >= curr_stop:
                continue

            start_x = int((curr_start / self.duration) * width)
            stop_x = int((curr_stop / self.duration) * width)
            seg_width = stop_x - start_x
            if seg_width < 1:
                seg_width = 1  # Minimum width to be visible
            
            # Pick color based on index
            color = self.colors[i % len(self.colors)]
            
            # Draw segment
            painter.fillRect(start_x, 0, seg_width, self.height(), color)


class ExportWorker(QObject):
    progress = pyqtSignal(int, int, str)
    finished = pyqtSignal(str, int, int, bool)

    def __init__(self, video_path, output_dir, segments):
        super().__init__()
        self.video_path = Path(video_path)
        self.output_dir = Path(output_dir)
        self.segments = segments
        self.cancel_requested = False

    def request_cancel(self):
        self.cancel_requested = True

    def run(self):
        total = len(self.segments)
        success_count = 0
        fail_count = 0

        for i, segment in enumerate(self.segments, start=1):
            if self.cancel_requested:
                break

            start_sec = max(0.0, segment["start"] / 1000.0)
            stop_sec = max(0.0, segment["stop"] / 1000.0)
            clip_duration = stop_sec - start_sec

            if clip_duration <= 0:
                fail_count += 1
                self.progress.emit(i, total, "Clip {} skipped (invalid segment)".format(i))
                continue

            output_file = self.output_dir / "{}_cut_{:02d}{}".format(self.video_path.stem, i, self.video_path.suffix)

            cmd = [
                "ffmpeg",
                "-hide_banner",
                "-loglevel", "error",
                "-nostdin",
                "-y",
                "-ss", str(start_sec),
                "-i", str(self.video_path),
                "-t", str(clip_duration),
                "-map", "0:v",
                "-map", "0:a?",
                # Use copy mode but enforce timestamp reset to fix playback issues
                "-c", "copy",
                "-avoid_negative_ts", "make_zero",
                # Move moov atom to start for better compatibility
                "-movflags", "+faststart",
                str(output_file)
            ]

            try:
                subprocess.run(
                    cmd,
                    check=True,
                    capture_output=True,
                    text=True,
                    timeout=900
                )
                success_count += 1
                self.progress.emit(i, total, "Clip {} exported".format(i))
            except subprocess.TimeoutExpired:
                fail_count += 1
                print("Timeout exporting clip {}".format(i), file=sys.stderr)
                self.progress.emit(i, total, "Clip {} failed (timeout)".format(i))
            except subprocess.CalledProcessError as err:
                fail_count += 1
                reason = (err.stderr or "").strip()
                print("FFmpeg Error for clip {}:".format(i), file=sys.stderr)
                print("Command: {}".format(" ".join(cmd)), file=sys.stderr)
                print("Error output:\n{}".format(reason), file=sys.stderr)
                
                reason_lines = reason.splitlines()
                reason_text = reason_lines[-1] if reason_lines else "ffmpeg error"
                self.progress.emit(i, total, "Clip {} failed ({})".format(i, reason_text))
            except Exception as err:
                fail_count += 1
                print("Exception exporting clip {}:".format(i), file=sys.stderr)
                traceback.print_exc()
                self.progress.emit(i, total, "Clip {} failed ({})".format(i, str(err)))

        self.finished.emit(str(self.output_dir), success_count, fail_count, self.cancel_requested)


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
        self._scrubbing = False  # True while the user is dragging the slider
        self._export_thread = None
        self._export_worker = None
        self._progress_dialog = None
        self._allow_close = False
        self._close_after_export = False
        self.immediate_export = False  # If True, skip showing UI and export immediately

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
        self.video_widget.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        layout.addWidget(self.video_widget, stretch=1)

        # Media player setup
        self.media_player = QMediaPlayer(None, QMediaPlayer.VideoSurface)
        self.media_player.setVideoOutput(self.video_widget)

        # ── Bottom control bar ──────────────────────────────────────────────
        control_bar = QWidget()
        control_bar.setFixedHeight(110)
        control_bar.setStyleSheet("background-color: #252525;")
        bar_layout = QVBoxLayout(control_bar)
        bar_layout.setContentsMargins(16, 10, 16, 10)
        bar_layout.setSpacing(8)

        # Timeline visualization bar
        self.segment_bar = SegmentBar()
        bar_layout.addWidget(self.segment_bar)

        # Time slider
        self.time_slider = QSlider(Qt.Horizontal)
        self.time_slider.sliderPressed.connect(self._on_slider_pressed)
        self.time_slider.sliderReleased.connect(self._on_slider_released)
        self.time_slider.sliderMoved.connect(self._on_slider_moved)
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
        self.play_button.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
        self.play_button.setIconSize(QSize(22, 22))
        self.play_button.setFixedSize(42, 42)
        self.play_button.clicked.connect(self.play_pause)
        row.addWidget(self.play_button)

        row.addStretch()

        # Segment buttons (right-aligned)
        self.start_button = QPushButton(u"\u25b6  Start")
        self.start_button.setObjectName("start_button")
        self.start_button.setFixedHeight(36)
        self.start_button.clicked.connect(self.on_start)
        row.addWidget(self.start_button)

        self.stop_and_start_button = QPushButton(u"\u23ed  Stop & Start")
        self.stop_and_start_button.setObjectName("stop_and_start_button")
        self.stop_and_start_button.setFixedHeight(36)
        self.stop_and_start_button.clicked.connect(self.on_stop_and_start)
        row.addWidget(self.stop_and_start_button)

        self.stop_button = QPushButton(u"\u25a0  Stop")
        self.stop_button.setObjectName("stop_button")
        self.stop_button.setFixedHeight(36)
        self.stop_button.clicked.connect(self.on_stop)
        row.addWidget(self.stop_button)

        row.addStretch()

        self.undo_button = QPushButton(u"\u21ba Undo")
        self.undo_button.setObjectName("undo_button")
        self.undo_button.setFixedHeight(36)
        self.undo_button.setToolTip("Remove the last segment")
        self.undo_button.clicked.connect(self.on_undo)
        row.addWidget(self.undo_button)

        bar_layout.addLayout(row)
        layout.addWidget(control_bar)

        # Initial state: not recording → show only Start
        self.is_recording = False
        self._update_button_state()

        # Connect media player signals
        self.media_player.positionChanged.connect(self.position_changed)
        self.media_player.durationChanged.connect(self.duration_changed)
        self.media_player.mediaStatusChanged.connect(self.on_media_status_changed)
        self.media_player.error.connect(self._handle_player_error)

    def _handle_player_error(self):
        self.play_button.setEnabled(False)
        error_string = self.media_player.errorString()
        error_code = self.media_player.error()
        err_msg = "Error (Code {}): {}".format(error_code, error_string)
        print(err_msg, file=sys.stderr)
        QMessageBox.critical(self, "Player Error", err_msg)
    
    def load_video(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Video File", "", "Video Files (*.mp4 *.avi *.mov *.mkv)"
        )
        
        if file_path:
            self.video_path = file_path
            self.media_player.setMedia(QMediaContent(QUrl.fromLocalFile(file_path)))

            # Reset state for new video
            self.segments = []
            self.is_recording = False
            self._update_button_state()
            
            # Check for existing splits
            json_path = Path(file_path).with_name(Path(file_path).stem + "_splits.json")
            if json_path.exists():
                reply = QMessageBox.question(
                    self, 
                    "Found Existing Splits", 
                    "A backup file with splits was found for this video.\nDo you want to load these splits?",
                    QMessageBox.Yes | QMessageBox.No
                )
                
                if reply == QMessageBox.Yes:
                    try:
                        with open(json_path, 'r') as f:
                            data = json.load(f)
                            
                        # Handle old format (list) vs new format (dict)
                        if isinstance(data, list):
                            self.segments = data
                        elif isinstance(data, dict):
                            self.segments = data.get("segments", [])
                        else:
                            self.segments = []

                        print("Loaded {} segments from backup.".format(len(self.segments)))

                        # Just load into editor and let user continue
                        self.immediate_export = False
                        
                        # Refresh UI state (button visibility, etc.)
                        if self.segments and self.segments[-1]["stop"] is None:
                            self.is_recording = True
                        self._update_button_state()
                        
                        # Set position to the end of the last recorded slice
                        if self.segments:
                            last_segment = self.segments[-1]
                            target_pos = 0
                            if last_segment.get("stop") is not None:
                                target_pos = last_segment["stop"]
                            elif last_segment.get("start") is not None:
                                target_pos = last_segment["start"]
                            
                            if target_pos > 0:
                                # We need to delay setting position slightly to ensure media is loaded
                                QTimer.singleShot(500, lambda: self.media_player.setPosition(target_pos))
                        
                        # Update bar with loaded segments
                        # Duration might not be known yet (media loading async), but set_data stores it.
                        # The subsequent durationChanged signal will also update it properly.
                        self.segment_bar.set_data(self.segments, self.media_player.duration())

                    except Exception as e:
                        QMessageBox.warning(self, "Load Error", "Failed to load splits: " + str(e))
    
    def play_pause(self):
        if self.media_player.state() == QMediaPlayer.PlayingState:
            self.media_player.pause()
            self.play_button.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
        else:
            self.media_player.play()
            self.play_button.setIcon(self.style().standardIcon(QStyle.SP_MediaPause))
    
    def _on_slider_pressed(self):
        self._scrubbing = True

    def _on_slider_released(self):
        self.media_player.setPosition(self.time_slider.value())
        self._scrubbing = False

    def _on_slider_moved(self, position):
        """Update the time label while dragging without seeking yet."""
        self.update_time_label(position, self.media_player.duration())

    def set_position(self, position):
        self.media_player.setPosition(position)
    
    def position_changed(self, position):
        if not self._scrubbing:
            self.time_slider.setValue(position)
        self.update_time_label(position, self.media_player.duration())
        
        # Update segment bar continuously to show current recording progress?
        # User said "already made splits", but visualizing the active one is nice.
        # But if we only want "already made", we don't need continuous updates, just on segment creation.
        # However, to be robust, let's update whenever segments or duration changes.
    
    def duration_changed(self, duration):
        self.time_slider.setRange(0, duration)
        self.segment_bar.set_data(self.segments, duration)
        self.segment_bar.update()

    def update_time_label(self, position, duration):
        self.time_label.setText(
            "{} / {}".format(self.format_time(position), self.format_time(duration))
        )

    def format_time(self, milliseconds):
        seconds = milliseconds // 1000
        minutes = seconds // 60
        seconds = seconds % 60
        return "{:02d}:{:02d}".format(minutes, seconds)

    def format_time_hms(self, milliseconds):
        """Format milliseconds as H:MM:SS for Google Sheets."""
        total_seconds = milliseconds // 1000
        hours = total_seconds // 3600
        minutes = (total_seconds % 3600) // 60
        seconds = total_seconds % 60
        return "{:d}:{:02d}:{:02d}".format(hours, minutes, seconds)
    
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
        print("Start point added at {}".format(self.format_time(current_time)))

    def on_stop(self):
        """Close the current segment — switches back to Start button."""
        current_time = self.media_player.position()
        self._close_current_segment(current_time)
        self.is_recording = False
        self._update_button_state()

        # Update visual bar
        self.segment_bar.set_data(self.segments, self.media_player.duration())

        # Save progress after every stop action
        self.save_state()

    def on_undo(self):
        """Removes the last segment (or cancels current recording)."""
        if not self.segments:
            print("Undo: No segments to undo.")
            return

        last_seg = self.segments.pop()
        
        if last_seg.get("stop") is None:
            # We were recording, so we cancel the start
            print("Undo: Cancelled current recording starting at {}".format(self.format_time(last_seg["start"])))
            self.is_recording = False
        else:
            # We removed a completed segment
            print("Undo: Removed segment starting at {}".format(
                self.format_time(last_seg["start"])
            ))

        self._update_button_state()
        self.segment_bar.set_data(self.segments, self.media_player.duration())
        self.save_state()


    def on_stop_and_start(self):
        """Close the current segment and immediately open a new one."""
        current_time = self.media_player.position()
        self._close_current_segment(current_time)
        self.segments.append({"start": current_time, "stop": None})
        print("New start point added at {}".format(self.format_time(current_time)))
        # Stays in recording state — buttons don't change

    def _close_current_segment(self, current_time):
        """Attach a stop time to the most recent open segment."""
        for segment in reversed(self.segments):
            if segment["stop"] is None:
                segment["stop"] = current_time
                print("Stop point added at {}".format(self.format_time(current_time)))
                return
        print("Warning: no open segment to close")
    
    def on_media_status_changed(self, status):
        """Once the video is loaded, read its frame rate and show the first frame."""
        print("Media status changed to: {}".format(status))
        if status == QMediaPlayer.InvalidMedia:
            print("Error: Invalid Media - Format might not be supported")
        
        if status == QMediaPlayer.LoadedMedia:
            fps = self.media_player.metaData("VideoFrameRate")
            if fps and fps > 0:
                self.frame_ms = int(1000 / fps)
                print("Detected frame rate: {:.2f} fps ({} ms/frame)".format(fps, self.frame_ms))
            else:
                print("Frame rate not found in metadata, using 30 fps default")

            # Play then immediately pause so the first frame is rendered
            self.media_player.play()
            self.media_player.pause()
            self.play_button.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))

    def keyPressEvent(self, event):
        """Left/right arrow keys step one frame backward/forward."""
        if event.key() == Qt.Key_Right:
            self.media_player.pause()
            self.play_button.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
            self.media_player.setPosition(
                min(self.media_player.position() + self.frame_ms, self.media_player.duration())
            )
        elif event.key() == Qt.Key_Left:
            self.media_player.pause()
            self.play_button.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
            self.media_player.setPosition(
                max(self.media_player.position() - self.frame_ms, 0)
            )
        else:
            super(VideoEditor, self).keyPressEvent(event)

    def save_state(self):
        if not self.video_path:
            return False
            
        data = {
            "segments": self.segments
        }
        
        try:
            json_path = self.video_path.with_name(self.video_path.stem + "_splits.json")
            with open(json_path, 'w') as f:
                json.dump(data, f, indent=4)
            print("Saved state to {}".format(json_path))
            return True
        except Exception as e:
            print("Failed to save state: {}".format(e), file=sys.stderr)
            return False
            
    def delete_state(self):
        if not self.video_path:
            return
        
        try:
            json_path = self.video_path.with_name(self.video_path.stem + "_splits.json")
            if json_path.exists():
                json_path.unlink()
                print("Deleted state file: {}".format(json_path))
        except Exception as e:
            print("Failed to delete state file: {}".format(e), file=sys.stderr)

    def closeEvent(self, event):
        """On close, ask user if they want to finish (export) or continue later (save)."""
        if self._allow_close:
            event.accept()
            return

        if self._export_thread and self._export_thread.isRunning():
            event.ignore()
            QMessageBox.information(self, "Export in progress", "Please wait for export to finish.")
            return

        # If no video loaded or no segments, just close
        if not self.video_path or not self.segments:
            event.accept()
            return
            
        # Ask user what to do
        reply = QMessageBox.question(
            self,
            "Video Editor",
            "Are you done with the video?\n\n"
            "Yes: Finish and export clips now. (Deletes saved progress)\n"
            "No: Save progress and continue later.",
            QMessageBox.Yes | QMessageBox.No | QMessageBox.Cancel
        )
        
        if reply == QMessageBox.Cancel:
            event.ignore()
            return
            
        if reply == QMessageBox.No:
            # Continue Later: Save progress and close
            self.save_state()
            self._allow_close = True
            event.accept()
            return
            
        # If Yes (Finished): Delete progress file and Export
        complete_segments = self._finalize_segments_for_export()
        
        if not complete_segments:
            # Maybe user clicked finish but there are no valid segments?
            event.accept()
            return

        event.ignore()
        self._close_after_export = True
        # State deletion moved to _on_export_finished to ensure it's only deleted after success
        self._start_export(complete_segments)


    def _finalize_segments_for_export(self):
        duration = max(0, self.media_player.duration())

        # Close any open segments so they can be exported
        for segment in self.segments:
            if segment["stop"] is None:
                segment["stop"] = duration
                print("Auto-closed open segment at {}".format(self.format_time(duration)))

        complete_segments = []
        for segment in self.segments:
            # Skip segments that are not valid duration
            if segment["stop"] is None:
                continue

            start = max(0, min(segment["start"], duration))
            stop = max(0, min(segment["stop"], duration))

            if stop > start:
                complete_segments.append({"start": start, "stop": stop})
            else:
                print("Skipping invalid segment: {} -> {}".format(start, stop))

        # Sort segments by start time (REMOVED: User requested order of creation)
        # complete_segments.sort(key=lambda s: s["start"])

        return complete_segments

    def _start_export(self, complete_segments):
        if shutil.which("ffmpeg") is None:
            QMessageBox.critical(self, "Missing ffmpeg", "ffmpeg was not found in PATH.")
            self._close_after_export = False
            return

        video_path = Path(self.video_path)
        output_dir = video_path.parent / "{}_clips".format(video_path.stem)

        try:
            output_dir.mkdir(exist_ok=True)
        except Exception as err:
            QMessageBox.critical(self, "Output folder error", "Could not create output folder:\n{}".format(err))
            self._close_after_export = False
            return

        # Save CSV report for Google Sheets
        self._save_csv_report(video_path, output_dir, complete_segments)

        if not self._has_enough_disk_space(video_path, output_dir, complete_segments):
            QMessageBox.warning(
                self,
                "Low disk space",
                "Not enough free disk space for export. Free up space and try again."
            )
            self._close_after_export = False
            return

        self.media_player.pause()
        self.play_button.setIcon(self.style().standardIcon(QStyle.SP_MediaPlay))
        
        # Hide the main editor window during export
        self.hide()

        # Use None as parent to ensure visibility independent of the hidden main window
        self._progress_dialog = QProgressDialog("Exporting clips...", "Cancel", 0, len(complete_segments), None)
        self._progress_dialog.setWindowTitle("Export")
        self._progress_dialog.setWindowModality(Qt.ApplicationModal)
        self._progress_dialog.setMinimumDuration(0)
        self._progress_dialog.setValue(0)
        self._progress_dialog.show()

        self._export_thread = QThread(self)
        self._export_worker = ExportWorker(video_path, output_dir, complete_segments)
        self._export_worker.moveToThread(self._export_thread)

        self._export_thread.started.connect(self._export_worker.run)
        self._export_worker.progress.connect(self._on_export_progress)
        self._export_worker.finished.connect(self._on_export_finished)
        self._progress_dialog.canceled.connect(self._export_worker.request_cancel)

        self._export_worker.finished.connect(self._export_thread.quit)
        self._export_worker.finished.connect(self._export_worker.deleteLater)
        self._export_thread.finished.connect(self._export_thread.deleteLater)

        self._export_thread.start()

    def _save_csv_report(self, video_path, output_dir, complete_segments):
        """Generates a CSV report for Google Sheets ingestion."""
        try:
            csv_path = output_dir / "{}_report.csv".format(video_path.stem)
            # Use newline='' for csv module to avoid extra blank lines on Windows
            with open(csv_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                # Header row
                writer.writerow(["Clip Name", "Start Time", "End Time"])
                
                for i, segment in enumerate(complete_segments, start=1):
                    # Replicate filename logic from ExportWorker
                    clip_name = "{}_cut_{:02d}{}".format(video_path.stem, i, video_path.suffix)
                    start_str = self.format_time_hms(segment["start"])
                    stop_str = self.format_time_hms(segment["stop"])
                    
                    writer.writerow([clip_name, start_str, stop_str])
            
            print("Saved CSV report to {}".format(csv_path))
        except Exception as e:
            print("Failed to save CSV report: {}".format(e), file=sys.stderr)

    def _has_enough_disk_space(self, video_path, output_dir, complete_segments):
        try:
            source_size = video_path.stat().st_size
            duration = max(1, self.media_player.duration())
            selected_ms = sum(max(0, s["stop"] - s["start"]) for s in complete_segments)
            estimated_bytes = int(source_size * (selected_ms / float(duration)) * 1.15)
            free_bytes = shutil.disk_usage(str(output_dir)).free
            return free_bytes >= estimated_bytes
        except Exception:
            # If estimate fails, avoid false blocking and attempt export.
            return True

    def _on_export_progress(self, current, total, message):
        if self._progress_dialog:
            self._progress_dialog.setLabelText("{} ({}/{})".format(message, current, total))
            self._progress_dialog.setValue(current)
        print(message)

    def _on_export_finished(self, output_dir, success_count, fail_count, canceled):
        if self._progress_dialog:
            self._progress_dialog.setValue(self._progress_dialog.maximum())
            self._progress_dialog.close()
            self._progress_dialog = None

        self._export_worker = None
        self._export_thread = None

        if canceled:
            QMessageBox.warning(
                self,
                "Export canceled",
                "Export canceled. {} clips exported, {} failed.".format(success_count, fail_count)
            )
            self._close_after_export = False
            return

        if fail_count > 0:
            QMessageBox.warning(
                self,
                "Export completed with errors",
                "{} clips exported, {} failed.\nSaved to:\n{}".format(success_count, fail_count, output_dir)
            )
        else:
            QMessageBox.information(
                self,
                "Export completed",
                "All clips exported successfully.\nSaved to:\n{}".format(output_dir)
            )
            # Only delete the state file if export was completely successful and we are closing
            if self._close_after_export:
                self.delete_state()

        if self._close_after_export:
            self._allow_close = True
            self.close()


def main():
    # Setup exception hook for verbose error printing
    def exception_hook(exctype, value, tb):
        print("CRITICAL: Uncaught exception:", file=sys.stderr)
        traceback.print_exception(exctype, value, tb)
        print("\n--- System Information ---", file=sys.stderr)
        print("Python:", sys.version, file=sys.stderr)
        print("Platform:", platform.platform(), file=sys.stderr)
        try:
            from PyQt5.QtCore import QT_VERSION_STR, PYQT_VERSION_STR
            print("Qt Version:", QT_VERSION_STR, file=sys.stderr)
            print("PyQt Version:", PYQT_VERSION_STR, file=sys.stderr)
        except ImportError:
            pass
        sys.exit(1)

    sys.excepthook = exception_hook

    app = QApplication(sys.argv)
    editor = VideoEditor()

    if not editor.immediate_export:
        editor.show()
    else:
        # If immediate export fails to start (e.g. video file selection canceled which shouldn't happen here
        # but if load_video failed or _start_export check failed), user might be stuck.
        # But QTimer fires after exec_ starts.
        pass

    sys.exit(app.exec_())


if __name__ == "__main__":
    main()
