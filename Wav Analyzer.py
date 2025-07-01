import sys
import os
import wave
import csv
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from PyQt5.QtWidgets import (QApplication, QMainWindow, QPushButton, QVBoxLayout, QHBoxLayout, 
                            QWidget, QLabel, QFileDialog, QMessageBox, QFrame, QTableWidget, 
                            QTableWidgetItem, QSplitter, QRadioButton, QButtonGroup, QSizePolicy, 
                            QGridLayout, QDialog, QTabWidget, QScrollArea, QTextBrowser, QInputDialog, QLineEdit)
from PyQt5.QtMultimedia import QSound
from PyQt5.QtCore import Qt, QRectF, QPropertyAnimation, QSize, pyqtSlot, QPoint, QSequentialAnimationGroup, QEasingCurve
from PyQt5.QtGui import QColor, QPalette, QFont, QDrag, QIcon, QLinearGradient, QRadialGradient, QPainter, QPen, QBrush, QPainterPath
from datetime import datetime
from scipy.io import wavfile


class AnimatedGradientWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMouseTracking(True)
        self.mouse_pos = QPoint(-100, -100)
        self.particles = []
        self.time = 0
        self.is_fullscreen = False
        
        # Initialize particles
        for _ in range(50):
            self.particles.append({
                'x': np.random.randint(0, 1000),
                'y': np.random.randint(0, 1000),
                'vx': np.random.randn() * 0.5,
                'vy': np.random.randn() * 0.5,
                'size': np.random.randint(3, 8)
            })
            
        # Start animation timer
        self.timer = self.startTimer(16)  # ~60 FPS
        
    def keyPressEvent(self, event):
        if event.key() == Qt.Key_F11:
            if self.is_fullscreen:
                self.window().showNormal()
            else:
                self.window().showFullScreen()
            self.is_fullscreen = not self.is_fullscreen
        super().keyPressEvent(event)
        
    def mouseMoveEvent(self, event):
        self.mouse_pos = event.pos()
        super().mouseMoveEvent(event)
        
    def timerEvent(self, event):
        self.time += 0.016
        
        # Update particle positions
        mouse_influence_radius = 150
        for p in self.particles:
            # Add mouse influence
            dx = self.mouse_pos.x() - p['x']
            dy = self.mouse_pos.y() - p['y']
            dist = np.sqrt(dx*dx + dy*dy)
            
            if dist < mouse_influence_radius:
                factor = (1 - dist/mouse_influence_radius) * 0.1
                p['vx'] += dx * factor
                p['vy'] += dy * factor
            
            # Update position
            p['x'] += p['vx']
            p['y'] += p['vy']
            
            # Add some random movement
            p['vx'] += (np.random.rand() - 0.5) * 0.1
            p['vy'] += (np.random.rand() - 0.5) * 0.1
            
            # Damping
            p['vx'] *= 0.99
            p['vy'] *= 0.99
            
            # Wrap around screen
            p['x'] = p['x'] % self.width()
            p['y'] = p['y'] % self.height()
        
        self.update()
        
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.Antialiasing)
        
        # Create animated gradient background
        gradient = QLinearGradient(0, 0, self.width(), self.height())
        t = (np.sin(self.time * 0.5) + 1) * 0.5  # Oscillate between 0 and 1
        
        # Green to black gradient with animation
        gradient.setColorAt(0, QColor(0, int(100 + 55 * t), 0))  # Dark green
        gradient.setColorAt(1, QColor(0, 0, 0))  # Black
        painter.fillRect(self.rect(), gradient)
        
        # Draw connecting lines between nearby particles
        painter.setPen(QPen(QColor(0, 255, 0, 30), 1))
        max_dist = 100
        
        for i, p1 in enumerate(self.particles):
            for p2 in self.particles[i+1:]:
                dx = p1['x'] - p2['x']
                dy = p1['y'] - p2['y']
                dist = np.sqrt(dx*dx + dy*dy)
                
                if dist < max_dist:
                    opacity = int(255 * (1 - dist/max_dist))
                    painter.setPen(QPen(QColor(0, 255, 0, opacity), 1))
                    painter.drawLine(int(p1['x']), int(p1['y']), int(p2['x']), int(p2['y']))
        
        # Draw particles
        for p in self.particles:
            color = QColor(0, 255, 0, 150)
            painter.setPen(Qt.NoPen)
            painter.setBrush(color)
            painter.drawEllipse(int(p['x'] - p['size']/2), int(p['y'] - p['size']/2), 
                              int(p['size']), int(p['size']))

class KatydidAnalysisApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Katydid Skip Analysis")
        self.setMinimumSize(800, 600)
        
        # Set window properties
        self.setWindowTitle("Katydid Call Analyzer")
        self.setGeometry(100, 100, 1400, 800)
        
        # Window state tracking
        self.is_fullscreen = False
        
        # Track original data for reset functionality
        self.original_wav_data = None
        
        # Initialize selection variables
        self.selection_start = None
        self.selection_end = None
        self.selection_ystart = None
        self.selection_yend = None
        self.selection_rect = None
        self.is_selecting = False
        
        # Set normal window flags (not frameless anymore)
        from PyQt5.QtCore import Qt
        self.setWindowFlags(Qt.Window)
        
        # Create close button that stays on top
        self.close_button = QPushButton("×", self)
        self.close_button.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        self.close_button.setStyleSheet("""
            QPushButton {
                color: red;
                background-color: transparent;
                font-size: 1.5em;
                border: none;
                font-weight: bold;
            }
            QPushButton:hover {
                color: #ff3333;
                background-color: rgba(255, 0, 0, 0.1);
            }
        """)
        self.close_button.raise_()
        self.close_button.raise_()
        self.close_button.clicked.connect(self.handle_close)
        
        # Initialize beep parameters for sound effect
        self.beep_duration = 100  # milliseconds
        
        # Start in fullscreen
        self.showFullScreen()
        
        # Initialize variables
        self.file_path = None
        self.file_loaded = False
        self.wav_data = None
        self.original_wav_data = None
        self.abs_data = None
        self.smoothed_data = None
        self.sample_rate = 44100  # Default sample rate
        self.sampwidth = 2  # Default sample width (16-bit)
        self.total_frames = 0
        self.chunk_size = 1000000  # Maximum chunk size to load at once
        self.current_chunk = None
        self.chunk_start = 0
        self.pulses = []
        self.skips = []
        self.threshold = 0.5
        self.abs_threshold = 0.5  # Absolute threshold for entire file
        self.rel_threshold = 0.5  # Relative threshold for current window
        self.using_absolute_threshold = True  # Flag to track which threshold is active
        self.view_start = 0
        self.view_range = 1000  # Initial view range in samples
        self.zoom_factor = 1.5
        self.file_type = None  # "wav" or "csv"
        self.short_pulse_avg = None
        self.long_pulse_avg = None
        self.file_loaded = False  # Track if file is loaded
        self.is_fullscreen = False  # Track fullscreen state
        
        # Initialize threshold labels
        self.threshold_label = QLabel("Threshold: 0.500")
        self.threshold_mode_label = QLabel("Mode: Absolute")
        
        # Set up the initial start screen
        self.setup_start_screen()

        # Set focus policy to accept keyboard focus
        from PyQt5.QtCore import Qt
        self.setFocusPolicy(Qt.StrongFocus)

        # Region selection variables
        self.region_selection_active = False
        self.region_selection_mode = None  # 'left' or 'right'
        self.region_left_pos = None
        self.region_right_pos = None
        self.region_left_line = None
        self.region_right_line = None
        self.region_rect = None

        # Initialize file queue for multiple file processing
        self.file_queue = []
        self.current_file_index = -1



    def setup_start_screen(self):
        if self.centralWidget():
            old_widget = self.centralWidget()
            old_widget.deleteLater()
        
        # Create main widget
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Create animated gradient background widget
        bg_widget = AnimatedGradientWidget()
        main_layout.addWidget(bg_widget)
        
        # Set as central widget
        self.setCentralWidget(main_widget)
        
        # Ensure close button is visible and on top
        self.close_button.setStyleSheet("""
            QPushButton {
                color: red;
                background-color: transparent;
                font-size: 32px;
                border: none;
                font-weight: bold;
            }
            QPushButton:hover {
                color: #ff3333;
                background-color: rgba(255, 0, 0, 0.1);
            }
        """)
        self.close_button.raise_()
        self.close_button.raise_()
        
        
        # Main layout for content
        layout = QVBoxLayout()
        bg_widget.setLayout(layout)
        
        # Title with animation
        title_label = QLabel("Katydid Call Analyzer")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("Arial", 36, QFont.Bold))
        title_label.setStyleSheet("""
            QLabel {
                color: #00ff00;
                margin: 20px;
                background-color: transparent;
                font-weight: bold;
            }
        """)
        
        # Create a pulsing animation for the title
        self.title_animation = QPropertyAnimation(title_label, b"minimumHeight")
        self.title_animation.setDuration(2000)
        self.title_animation.setStartValue(80)
        self.title_animation.setEndValue(90)
        self.title_animation.setLoopCount(-1)  # Infinite loop
        self.title_animation.setEasingCurve(QEasingCurve.InOutQuad)
        
        # Create a sequential animation group to go back and forth
        self.title_animation_group = QSequentialAnimationGroup()
        self.title_animation_group.addAnimation(self.title_animation)
        
        # Create the reverse animation
        reverse_animation = QPropertyAnimation(title_label, b"minimumHeight")
        reverse_animation.setDuration(2000)
        reverse_animation.setStartValue(90)
        reverse_animation.setEndValue(80)
        reverse_animation.setEasingCurve(QEasingCurve.InOutQuad)
        
        self.title_animation_group.addAnimation(reverse_animation)
        self.title_animation_group.start()
        
        # Description
        description = QLabel("A tool for analyzing katydid calls from WAV files")
        description.setAlignment(Qt.AlignCenter)
        description.setFont(QFont("Arial", 16))
        description.setStyleSheet("""
            QLabel {
                color: white;
                margin: 10px;
                background-color: transparent;
            }
        """)
        
        # Start button with modern design
        start_button = QPushButton("Start Analysis")
        start_button.setFixedSize(240, 60)
        start_button.setFont(QFont("Arial", 16))
        start_button.setStyleSheet("""
            QPushButton {
                background-color: rgba(0, 40, 0, 0.8);
                color: #00ff00;
                border: 2px solid #00ff00;
                border-radius: 30px;
                padding: 15px;
            }
            QPushButton:hover {
                background-color: rgba(0, 60, 0, 0.9);
                color: white;
            }
            QPushButton:pressed {
                background-color: rgba(0, 80, 0, 1.0);
                border: 2px solid #ffffff;
            }
        """)
        
        # Connect button to transition function with developer notes popup first
        start_button.clicked.connect(self.show_developer_notes)
        
        # Layout everything
        layout.addStretch()
        layout.addWidget(title_label)
        layout.addSpacing(20)
        layout.addWidget(description)
        layout.addSpacing(50)
        layout.addWidget(start_button, alignment=Qt.AlignCenter)
        layout.addStretch()
        
        # Play startup sound
        self.play_startup_sound()
    
    def show_developer_notes(self):
        """Show developer notes popup with recent updates and features."""
        notes_dialog = QDialog(self)
        notes_dialog.setWindowTitle("Developer Notes")
        notes_dialog.setMinimumSize(700, 600)  # Increased size
        notes_dialog.setStyleSheet("""
            QDialog {
                background-color: #1a1a1a;
                color: #ffffff;
                border: 2px solid #00ff00;
                border-radius: 10px;
            }
        """)
        
        # Create layout
        layout = QVBoxLayout(notes_dialog)
        
        # Title
        title = QLabel("Developer Notes - Updates")
        title.setFont(QFont("Arial", 18, QFont.Bold))  # Increased font size
        title.setStyleSheet("color: #00ff00; margin-bottom: 15px;")
        title.setAlignment(Qt.AlignCenter)
        
        # Notes content
        notes_content = QTextBrowser()
        notes_content.setStyleSheet("""
            QTextBrowser {
                background-color: #2a2a2a;
                color: #ffffff;
                border: 1px solid #444444;
                border-radius: 5px;
                padding: 10px;
                font-size: 12pt;  /* Increased font size */
            }
        """)
        
        # Set the HTML content with bullet points and keyboard controls
        notes_content.setHtml("""
        <html>
        <body style="font-family: Arial; font-size: 12pt;">
        <h2 style="color: #00ff00;">Latest Updates</h2>
        <ul>
            <li>Adding CSV analysis soon</li>
            <li>Recently fixed save function</li>
            <li>All functionalities should be working</li>
            <li>Now you can select multiple WAV files to analyze simply by dragging and dropping a folder or highlighting multiple waveforms</li>
        </ul>
        
        <h2 style="color: #00ff00;">Keyboard Controls</h2>
        <table border="0" cellspacing="5" cellpadding="5" width="100%">
            <tr>
                <td width="30%"><b>Navigation:</b></td>
                <td width="70%"></td>
            </tr>
            <tr>
                <td>W</td>
                <td>Zoom in</td>
            </tr>
            <tr>
                <td>S</td>
                <td>Zoom out</td>
            </tr>
            <tr>
                <td>A</td>
                <td>Move left</td>
            </tr>
            <tr>
                <td>D</td>
                <td>Move right</td>
            </tr>
            <tr>
                <td><b>Analysis:</b></td>
                <td></td>
            </tr>
            <tr>
                <td>Y</td>
                <td>Detect pulses</td>
            </tr>
            <tr>
                <td>T</td>
                <td>Analyze periods</td>
            </tr>
            <tr>
                <td>R</td>
                <td>Invert values</td>
            </tr>
            <tr>
                <td>G</td>
                <td>Smooth signal</td>
            </tr>
            <tr>
                <td>/</td>
                <td>Start region selection</td>
            </tr>
            <tr>
                <td><b>Threshold:</b></td>
                <td></td>
            </tr>
            <tr>
                <td>Up Arrow</td>
                <td>Increase threshold</td>
            </tr>
            <tr>
                <td>Down Arrow</td>
                <td>Decrease threshold</td>
            </tr>
            <tr>
                <td><b>Other:</b></td>
                <td></td>
            </tr>
            <tr>
                <td>=</td>
                <td>Save results with WAV file</td>
            </tr>
            <tr>
                <td>F11</td>
                <td>Toggle fullscreen</td>
            </tr>
            <tr>
                <td>Escape</td>
                <td>Exit fullscreen</td>
            </tr>
        </table>
        
        <h2 style="color: #00ff00;">Multiple File Processing</h2>
        <p>You can now process multiple WAV files in sequence:</p>
        <ul>
            <li>Select multiple files using the file dialog</li>
            <li>Drag and drop a folder containing WAV files</li>
            <li>After saving results for one file, the app will automatically load the next file</li>
        </ul>
        
        <p>For questions or feedback, please contact David Nguyen at davidminhnguyen2@gmail.com.</p>
        </body>
        </html>
        """)
        
        # Continue button
        continue_button = QPushButton("Continue to Analysis")
        continue_button.setFixedSize(200, 40)
        continue_button.setStyleSheet("""
            QPushButton {
                background-color: rgba(0, 40, 0, 0.8);
                color: #00ff00;
                border: 2px solid #00ff00;
                border-radius: 15px;
                padding: 8px;
            }
            QPushButton:hover {
                background-color: rgba(0, 60, 0, 0.9);
                color: white;
            }
            QPushButton:pressed {
                background-color: rgba(0, 80, 0, 1.0);
                border: 2px solid #ffffff;
            }
        """)
        
        # Connect button to close dialog and continue
        continue_button.clicked.connect(lambda: self.continue_to_analysis(notes_dialog))
        
        # Add widgets to layout
        layout.addWidget(title)
        layout.addWidget(notes_content)
        layout.addWidget(continue_button, alignment=Qt.AlignCenter)
        
        # Show the dialog
        notes_dialog.exec_()
    
    def continue_to_analysis(self, dialog):
        """Close the developer notes dialog and continue to analysis."""
        dialog.accept()
        self.transition_to_pulse_selection()
    
    def handle_close(self):
        reply = QMessageBox.question(
            self, 'Exit Application',
            'Are you sure you want to exit?',
            QMessageBox.Yes | QMessageBox.No,
            QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.close()
    
    def keyPressEvent(self, event):
        """Handle keyboard events."""
        from PyQt5.QtCore import Qt
        
        # Only process if a file is loaded
        if not hasattr(self, 'file_path') or not self.file_path:
            super().keyPressEvent(event)  # Pass to parent for default handling
            return
            
        key = event.key()
        
        # Navigation controls
        if key == Qt.Key_A:  # Move left
            self.move_view(-1)
        elif key == Qt.Key_D:  # Move right
            self.move_view(1)
        elif key == Qt.Key_W:  # Zoom in
            self.zoom_view(0.5)  # Zoom in by factor of 2
        elif key == Qt.Key_S:  # Zoom out
            self.zoom_view(2.0)  # Zoom out by factor of 2
        
        # Selection controls
        elif key == Qt.Key_O:  # Add pulse at selection
            self.add_manual_pulse()
        elif key == Qt.Key_P:  # Delete pulses in selection
            self.delete_selected_pulses()
        
        # Processing controls
        elif key == Qt.Key_R:  # Invert values
            self.invert_values()
        elif key == Qt.Key_G:  # Smooth signal
            self.apply_smoothing()
        elif key == Qt.Key_Y:  # Detect pulses
            self.detect_pulses()
        elif key == Qt.Key_T:  # Analyze pulse periods
            self.analyze_pulse_periods()
        
        # Threshold controls
        elif key == Qt.Key_Up:  # Increase threshold
            if self.using_absolute_threshold:
                self.abs_threshold = min(1.0, self.abs_threshold + 0.025)
            else:
                self.rel_threshold = min(1.0, self.rel_threshold + 0.025)
            self.update_plot()
        elif key == Qt.Key_Down:  # Decrease threshold
            if self.using_absolute_threshold:
                self.abs_threshold = max(0.0, self.abs_threshold - 0.025)
            else:
                self.rel_threshold = max(0.0, self.rel_threshold - 0.025)
            self.update_plot()
        elif key == Qt.Key_BracketLeft or key == Qt.Key_BracketRight:  # Toggle threshold mode
            self.using_absolute_threshold = not self.using_absolute_threshold
            self.threshold = self.abs_threshold if self.using_absolute_threshold else self.rel_threshold
            self.threshold_mode_label.setText(f"Mode: {'Absolute' if self.using_absolute_threshold else 'Relative'}")
            # For relative threshold, update based on visible data
            if not self.using_absolute_threshold and self.smoothed_data is not None:
                visible_data = self.smoothed_data[self.view_start:min(self.view_start + self.view_range, len(self.smoothed_data))]
                self.threshold = self.rel_threshold * np.max(visible_data) if len(visible_data) > 0 else self.rel_threshold
            self.update_plot()

        # Save results
        elif key == Qt.Key_Equal:  # Save results with WAV file
            self.save_results_with_wav()

        # F11 and Escape key handling
        elif key == Qt.Key_F11:
            if self.isFullScreen():
                self.showNormal()
            else:
                self.showFullScreen()
        elif key == Qt.Key_Escape and self.isFullScreen():
            self.showNormal()
        else:
            super().keyPressEvent(event)
        
    def resizeEvent(self, event):
        super().resizeEvent(event)
        # Update close button position
        if hasattr(self, 'close_button'):
            self.close_button.move(self.width() - 45, 5)
    
    def play_startup_sound(self):
        try:
            import winsound
            # Play a sequence of ascending notes for a modern startup sound
            freqs = [523, 659, 784, 988]  # C5, E5, G5, B5
            for freq in freqs:
                winsound.Beep(freq, 50)
        except:
            pass

    def transition_to_pulse_selection(self):
        # Play startup sound
        self.play_startup_sound()
        
        # Store current widget
        old_widget = self.centralWidget()
        
        # Create new widget
        self.show_pulse_selection()
        new_widget = self.centralWidget()
        new_widget.setWindowOpacity(0.0)
        
        # Setup fade out animation
        fade_out = QPropertyAnimation(old_widget, b'windowOpacity')
        fade_out.setDuration(300)
        fade_out.setStartValue(1.0)
        fade_out.setEndValue(0.0)
        
        # Setup fade in animation
        fade_in = QPropertyAnimation(new_widget, b'windowOpacity')
        fade_in.setDuration(300)
        fade_in.setStartValue(0.0)
        fade_in.setEndValue(1.0)
        
        # Run animations in sequence
        sequence = QSequentialAnimationGroup()
        sequence.addAnimation(fade_out)
        sequence.addAnimation(fade_in)
        sequence.finished.connect(lambda: old_widget.deleteLater())
        sequence.start()

    def show_pulse_selection(self):
        # Create main widget
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Create animated gradient background widget
        bg_widget = AnimatedGradientWidget()
        main_layout.addWidget(bg_widget)
        
        # Set as central widget
        self.setCentralWidget(main_widget)
        
        # Ensure close button is visible and on top
        self.close_button.setStyleSheet("""
            QPushButton {
                color: red;
                background-color: transparent;
                font-size: 32px;
                border: none;
                font-weight: bold;
            }
            QPushButton:hover {
                color: #ff3333;
                background-color: rgba(255, 0, 0, 0.1);
            }
        """)
        self.close_button.raise_()
        self.close_button.raise_()
        
        
        # Main layout for content
        layout = QVBoxLayout()
        bg_widget.setLayout(layout)
        
        # Title
        title_label = QLabel("Katydid Call Analyzer")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("Arial", 32, QFont.Bold))
        title_label.setStyleSheet("""
            QLabel {
                color: #00ff00;
                margin: 20px;
                background-color: transparent;
                font-weight: bold;
            }
        """)
        
        # Description
        description = QLabel("Click the button below to start analyzing WAV files")
        description.setAlignment(Qt.AlignCenter)
        description.setFont(QFont("Arial", 16))
        description.setStyleSheet("""
            QLabel {
                color: white;
                margin: 10px;
                background-color: transparent;
            }
        """)
        
        # Continue button
        continue_button = QPushButton("Start Analysis")
        continue_button.setFixedSize(200, 60)
        continue_button.setFont(QFont("Arial", 16))
        continue_button.setStyleSheet("""
            QPushButton {
                background-color: rgba(0, 40, 0, 0.8);
                color: #00ff00;
                border: 2px solid #00ff00;
                border-radius: 30px;
                padding: 10px;
            }
            QPushButton:enabled:hover {
                background-color: rgba(0, 60, 0, 0.9);
                color: white;
            }
            QPushButton:pressed {
                background-color: rgba(0, 80, 0, 1.0);
                border: 2px solid #ffffff;
            }
        """)
        
        # Set file type to WAV by default
        self.file_type = "wav"
        
        # Connect button to analysis interface
        continue_button.clicked.connect(self.setup_analysis_interface)
        
        # Layout everything
        layout.addStretch()
        layout.addWidget(title_label)
        layout.addSpacing(20)
        layout.addWidget(description)
        layout.addSpacing(50)
        layout.addWidget(continue_button, alignment=Qt.AlignCenter)
        layout.addStretch()
    
    def setup_analysis_interface(self):
        # Clear any existing layout
        if self.centralWidget():
            self.centralWidget().deleteLater()
        
        # Create main widget
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)
        
        # Create content widget
        central_widget = QWidget()
        main_layout.addWidget(central_widget)
        
        # Set as central widget
        self.setCentralWidget(main_widget)
        
        # Ensure close button is visible and on top
        self.close_button.setStyleSheet("""
            QPushButton {
                color: red;
                background-color: transparent;
                font-size: 32px;
                border: none;
                font-weight: bold;
            }
            QPushButton:hover {
                color: #ff3333;
                background-color: rgba(255, 0, 0, 0.1);
            }
        """)
        self.close_button.raise_()
        self.close_button.raise_()
    
        
        main_layout = QVBoxLayout()
        central_widget.setLayout(main_layout)
        central_widget.setStyleSheet("""
            QWidget {
                background-color: qlineargradient(spread:pad, x1:0, y1:0, x2:1, y2:1, 
                                      stop:0 #86c67c, stop:1 #4CAF50);
            }
        """)
        
        # Title with file type
        file_type_text = "WAV" if hasattr(self, 'file_type') and self.file_type == "wav" else "CSV"
        title_label = QLabel(f"Katydid {file_type_text} Analysis")
        title_label.setAlignment(Qt.AlignCenter)
        title_label.setFont(QFont("Arial", 16, QFont.Bold))
        title_label.setStyleSheet("color: white; background-color: transparent;")
        main_layout.addWidget(title_label)
        
        # File drop area
        self.file_drop_frame = QFrame()
        self.file_drop_frame.setFrameShape(QFrame.StyledPanel)
        self.file_drop_frame.setStyleSheet("""
            QFrame {
                background-color: rgba(255, 255, 255, 0.8);
                border: 2px dashed #aaa;
                border-radius: 10px;
                min-height: 100px;
            }
            QFrame:hover {
                background-color: rgba(255, 255, 255, 0.9);
                border: 2px dashed #666;
            }
        """)
        file_drop_layout = QVBoxLayout(self.file_drop_frame)
        
        self.drop_label = QLabel("Drag and drop a .wav file here\nor click to browse")
        self.drop_label.setAlignment(Qt.AlignCenter)
        self.drop_label.setFont(QFont("Arial", 12))
        file_drop_layout.addWidget(self.drop_label)
        
        self.browse_button = QPushButton("Browse Files")
        self.browse_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 5px;
                padding: 8px;
                max-width: 150px;
            }
            QPushButton:hover {
                background-color: #45a049;
                border: 1px solid white;
            }
        """)
        file_drop_layout.addWidget(self.browse_button, alignment=Qt.AlignCenter)
        
        # Connect browse button to file dialog
        self.browse_button.clicked.connect(self.open_file_dialog)
        
        main_layout.addWidget(self.file_drop_frame)
        
        # Matplotlib figure for waveform display (initially hidden)
        self.figure = Figure(tight_layout=True)
        self.canvas = FigureCanvas(self.figure)
        self.canvas.setMinimumHeight(400)  # Set minimum height
        self.canvas.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)  # Make it expand
        self.canvas.setVisible(False)
        self.ax = self.figure.add_subplot(111)
        
        # Enhanced canvas event connections with more robust capture
        self.canvas.mpl_connect('button_press_event', self.on_mouse_press)
        self.canvas.mpl_connect('button_release_event', self.on_mouse_release)
        self.canvas.mpl_connect('motion_notify_event', self.on_mouse_move)
        
        # Set interactive mode for more responsive updates
        self.figure.set_tight_layout(True)
        
        main_layout.addWidget(self.canvas)
        
        # Controls panel (initially hidden)
        self.controls_widget = QWidget()
        self.controls_widget.setVisible(False)
        self.controls_widget.setStyleSheet("background-color: transparent;")
        controls_layout = QHBoxLayout(self.controls_widget)
        
        # Left Panel (Navigation & Processing)
        left_frame = QFrame()
        left_frame.setFrameShape(QFrame.StyledPanel)
        left_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border-radius: 12px;
                padding: 20px;
                border: 1px solid #dee2e6;
            }
        """)
        left_layout = QVBoxLayout(left_frame)
        left_layout.setSpacing(15)
        
        # Navigation Controls Section
        nav_label = QLabel("Navigation Controls")
        nav_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        nav_label.setStyleSheet("""
            QLabel {
                color: #212529;
                border-bottom: 2px solid #6c757d;
                padding-bottom: 10px;
                margin-bottom: 15px;
            }
        """)
        left_layout.addWidget(nav_label)
        
        # Create button grid layout for navigation controls
        nav_buttons = QGridLayout()
        nav_buttons.setSpacing(12)
        
        # Button style
        button_style = """
            QPushButton {
                background-color: #f8f9fa;
                color: #212529;
                border: 1px solid #ced4da;
                border-radius: 8px;
                padding: 10px 16px;
                font-size: 13px;
                font-weight: 500;
                min-width: 120px;
                text-align: center;
            }
            QPushButton:hover {
                background-color: #e9ecef;
                border-color: #adb5bd;
            }
            QPushButton:pressed {
                background-color: #dee2e6;
                border-color: #6c757d;
            }
        """
        
        # Create navigation buttons
        left_btn = QPushButton("← Left (A)")
        left_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        left_btn.setStyleSheet(button_style)
        left_btn.clicked.connect(lambda: self.move_view(-1))
        
        right_btn = QPushButton("Right (D) →")
        right_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        right_btn.setStyleSheet(button_style)
        right_btn.clicked.connect(lambda: self.move_view(1))
        
        zoom_in_btn = QPushButton("Zoom In (W)")
        zoom_in_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        zoom_in_btn.setStyleSheet(button_style)
        zoom_in_btn.clicked.connect(lambda: self.zoom_view(0.5))
        
        zoom_out_btn = QPushButton("Zoom Out (S)")
        zoom_out_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        zoom_out_btn.setStyleSheet(button_style)
        zoom_out_btn.clicked.connect(lambda: self.zoom_view(2.0))
        
        # Add buttons to grid
        nav_buttons.addWidget(left_btn, 0, 0)
        nav_buttons.addWidget(right_btn, 0, 1)
        nav_buttons.addWidget(zoom_in_btn, 1, 0)
        nav_buttons.addWidget(zoom_out_btn, 1, 1)
        
        left_layout.addLayout(nav_buttons)
        
        # Processing Controls Section
        proc_label = QLabel("Processing Controls")
        proc_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        proc_label.setStyleSheet("""
            QLabel {
                color: #212529;
                border-bottom: 2px solid #6c757d;
                padding-bottom: 10px;
                margin-bottom: 15px;
                margin-top: 20px;
            }
        """)
        left_layout.addWidget(proc_label)
        
        # Create button grid layout for processing controls
        proc_buttons = QGridLayout()
        proc_buttons.setSpacing(12)
        
        # Create processing buttons
        detect_btn = QPushButton("Detect Pulses (Y)")
        detect_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        detect_btn.setStyleSheet(button_style)
        detect_btn.clicked.connect(self.detect_pulses)
        
        analyze_periods_btn = QPushButton("Analyze Periods (T)")
        analyze_periods_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        analyze_periods_btn.setStyleSheet(button_style)
        analyze_periods_btn.clicked.connect(self.analyze_pulse_periods)
        
        invert_btn = QPushButton("Invert Values (R)")
        invert_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        invert_btn.setStyleSheet(button_style)
        invert_btn.clicked.connect(self.invert_values)
        
        smooth_btn = QPushButton("Smooth Signal (G)")
        smooth_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        smooth_btn.setStyleSheet(button_style)
        smooth_btn.clicked.connect(self.apply_smoothing)
        
        # Create reset and help buttons
        reset_btn = QPushButton("Reset (Clear All)")
        reset_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        reset_btn.setStyleSheet(button_style)
        reset_btn.clicked.connect(self.reset_application)
        
        help_btn = QPushButton("Show Controls")
        help_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        help_btn.setStyleSheet(button_style)
        help_btn.clicked.connect(self.show_help)
        
        # Add buttons to grid
        proc_buttons.addWidget(detect_btn, 0, 0)
        proc_buttons.addWidget(analyze_periods_btn, 0, 1)
        proc_buttons.addWidget(invert_btn, 1, 0)
        proc_buttons.addWidget(smooth_btn, 1, 1)
        proc_buttons.addWidget(reset_btn, 2, 0)
        proc_buttons.addWidget(help_btn, 2, 1)
        
        left_layout.addLayout(proc_buttons)
        
        # Right Panel (Threshold & Selection)
        right_frame = QFrame()
        right_frame.setFrameShape(QFrame.StyledPanel)
        right_frame.setStyleSheet("""
            QFrame {
                background-color: #f8f9fa;
                border-radius: 12px;
                padding: 20px;
                border: 1px solid #dee2e6;
            }
        """)
        right_layout = QVBoxLayout(right_frame)
        right_layout.setSpacing(15)
        
        # Add threshold controls information
        threshold_label = QLabel("Threshold Controls")
        threshold_label.setStyleSheet("""
            QLabel {
                color: #212529;
                border-bottom: 2px solid #6c757d;
                padding-bottom: 10px;
                margin-bottom: 15px;
                font-size: 1.2em;
                font-weight: bold;
            }
        """)
        right_layout.addWidget(threshold_label)
        
        # Add threshold controls info
        threshold_info = QLabel("↑/↓: Adjust threshold\n[/]: Switch absolute/relative mode")
        threshold_info.setAlignment(Qt.AlignCenter)
        threshold_info.setStyleSheet("""
            QLabel {
                color: #495057;
                background-color: #e9ecef;
                padding: 1em;
                border-radius: 6px;
                margin-bottom: 15px;
                font-size: 1em;
            }
        """)
        right_layout.addWidget(threshold_info)
        
        # Selection Controls Section
        selection_label = QLabel("Selection Controls")
        selection_label.setStyleSheet("""
            QLabel {
                color: #212529;
                border-bottom: 2px solid #6c757d;
                padding-bottom: 10px;
                margin-bottom: 15px;
                margin-top: 20px;
                font-size: 1.2em;
                font-weight: bold;
            }
        """)
        right_layout.addWidget(selection_label)
        
        # Create button grid layout for selection controls
        selection_buttons = QGridLayout()
        selection_buttons.setSpacing(12)
        
        # Create selection buttons
        add_pulse_btn = QPushButton("Add Pulse (O)")
        add_pulse_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        add_pulse_btn.setStyleSheet(button_style)
        add_pulse_btn.clicked.connect(self.add_manual_pulse)
        
        delete_pulse_btn = QPushButton("Delete Pulse (P)")
        delete_pulse_btn.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        delete_pulse_btn.setStyleSheet(button_style)
        delete_pulse_btn.clicked.connect(self.delete_selected_pulses)
        
        # Add buttons to grid
        selection_buttons.addWidget(add_pulse_btn, 0, 0)
        selection_buttons.addWidget(delete_pulse_btn, 0, 1)
        
        right_layout.addLayout(selection_buttons)
        
        # Add save button with distinct style
        save_button = QPushButton("Save Results")
        save_button.setSizePolicy(QSizePolicy.Minimum, QSizePolicy.Minimum)
        save_button.clicked.connect(self.save_results)
        save_button.setStyleSheet("""
            QPushButton {
                background-color: #4CAF50;
                color: white;
                border-radius: 8px;
                padding: 1em 1.5em;
                font-weight: bold;
                font-size: 1em;
                margin-top: 20px;
                border: none;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #45a049;
            }
            QPushButton:pressed {
                background-color: #3d8b40;
            }
        """)
        right_layout.addWidget(save_button)
        
        # Add both frames to the controls layout
        controls_layout.addWidget(left_frame)
        controls_layout.addWidget(right_frame)
        
        main_layout.addWidget(self.controls_widget)
        
        # Add threshold controls and info
        threshold_frame = QFrame()
        threshold_frame.setStyleSheet("""
            QFrame {
                background-color: rgba(255, 255, 255, 0.8);
                border-radius: 5px;
                padding: 5px;
            }
        """)
        threshold_layout = QVBoxLayout(threshold_frame)
        
        # Threshold mode toggle
        self.threshold_mode_label = QLabel("Mode: Absolute")
        self.threshold_mode_label.setFont(QFont("Arial", 10))
        self.threshold_mode_label.setStyleSheet("color: #4CAF50;")
        threshold_layout.addWidget(self.threshold_mode_label)
        
        # Threshold value display
        self.threshold_label = QLabel("Threshold: 0.0")
        self.threshold_label.setFont(QFont("Arial", 10))
        self.threshold_label.setStyleSheet("color: #4CAF50;")
        threshold_layout.addWidget(self.threshold_label)
        
        controls_layout.addWidget(threshold_frame)
        
        # Time display
        self.time_label = QLabel("")
        self.time_label.setFont(QFont("Arial", 10))
        self.time_label.setStyleSheet("color: #4CAF50;")
        controls_layout.addWidget(self.time_label)
        
        # Status bar
        self.statusBar().showMessage("Ready")
    
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = event.mimeData().urls()
        self.file_queue = []
        
        # Process all dropped files/folders
        for url in urls:
            path = url.toLocalFile()
            
            # Check if it's a directory
            if os.path.isdir(path):
                # Get all WAV files in the directory
                wav_files = [os.path.join(path, f) for f in os.listdir(path) 
                            if f.lower().endswith('.wav') and os.path.isfile(os.path.join(path, f))]
                # Sort alphabetically
                wav_files.sort()
                # Add to queue
                self.file_queue.extend(wav_files)
            # Check if it's a WAV file
            elif os.path.isfile(path) and path.lower().endswith('.wav'):
                self.file_queue.append(path)
        
        # If we have files in the queue, start processing
        if self.file_queue:
            self.current_file_index = 0
            self.load_wav_file(self.file_queue[self.current_file_index])
            if len(self.file_queue) > 1:
                QMessageBox.information(self, "Multiple Files", 
                                      f"Loaded {len(self.file_queue)} files. Processing will continue automatically after saving results.")
        else:
            QMessageBox.warning(self, "Invalid Files", "No valid WAV files found. Please drop WAV files or a folder containing WAV files.")
    
    def open_file_dialog(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, "Open WAV Files", "", "WAV Files (*.wav)")
        if file_paths:
            # Add files to the queue
            self.file_queue = file_paths
            self.current_file_index = 0
            # Load the first file
            self.load_wav_file(self.file_queue[self.current_file_index])
    
    def load_wav_file(self, file_path):
        try:
            # Ensure file exists and is readable
            if not os.path.exists(file_path):
                raise FileNotFoundError(f"File not found: {file_path}")
            
            # Load the entire WAV file at once using scipy.io.wavfile
            self.sample_rate, wav_data = wavfile.read(file_path)
            
            # Convert to float in range [-1, 1] for consistent processing
            if wav_data.dtype == np.int16:
                wav_data = wav_data.astype(np.float32) / 32768.0
            elif wav_data.dtype == np.int32:
                wav_data = wav_data.astype(np.float32) / 2147483648.0
            elif wav_data.dtype == np.uint8:
                wav_data = (wav_data.astype(np.float32) - 128) / 128.0
                
            # If stereo, convert to mono by averaging channels
            if len(wav_data.shape) > 1 and wav_data.shape[1] > 1:
                wav_data = np.mean(wav_data, axis=1)
                
            # Store total frames
            self.total_frames = len(wav_data)
            
            # Store file path
            self.file_path = file_path
            
            # Calculate file size in MB
            file_size_mb = os.path.getsize(file_path) / (1024 * 1024)
            
            # Store the entire audio data
            self.wav_data = wav_data
            
            # Keep a copy of the original data for reset functionality
            self.original_wav_data = wav_data.copy()
            
            # Reset processing variables
            self.abs_data = None
            self.smoothed_data = None
            self.pulses = []
            self.skips = []
            
            # Track number of inversions
            self.inversion_count = 0
            
            # Initialize processing variables
            self.abs_threshold = 0.5
            self.rel_threshold = 0.5
            self.threshold = self.abs_threshold
            self.using_absolute_threshold = True
            
            # Configure the view
            self.view_start = 0
            self.view_range = min(self.sample_rate * 2, self.total_frames)  # View first 2 seconds
            
            # Show the waveform and controls
            self.canvas.setVisible(True)
            self.controls_widget.setVisible(True)
            
            # Update the file drop area
            self.drop_label.setText(f"File loaded: {os.path.basename(file_path)}")
            self.browse_button.setVisible(False)
            self.file_loaded = True
            
            # Update the plot
            self.update_plot()
            
            # Display file information
            duration = self.total_frames / self.sample_rate
            time_s = duration
            time_m = int(time_s // 60)
            time_s = time_s % 60
            
            QMessageBox.information(self, "File Loaded", 
                f"File: {os.path.basename(file_path)}\n"
                f"Duration: {time_m}m {time_s:.2f}s\n"
                f"Sample Rate: {self.sample_rate} Hz\n"
                f"File Size: {file_size_mb:.1f}MB")
            
            # Set focus to the central widget for keyboard shortcuts
            self.centralWidget().setFocus()
            
        except KeyboardInterrupt:
            QMessageBox.warning(self, "Operation Cancelled", "File loading was cancelled by user")
            return
        except Exception as e:
            QMessageBox.critical(self, "Error Loading File", f"Could not load the WAV file: {str(e)}")
            return
    
    def load_chunk(self, start_frame):
        """Load a chunk of audio data starting from start_frame"""
        try:
            if not hasattr(self, 'file_path') or not self.file_path:
                return None
            
            # Validate start frame
            if start_frame >= self.total_frames:
                return None
            
            # Calculate frames to read
            frames_to_read = min(self.chunk_size, self.total_frames - start_frame)
            
            # Load data in smaller chunks to avoid memory issues
            chunk_size = 1024  # Load in 1KB chunks
            num_chunks = frames_to_read // chunk_size
            remaining = frames_to_read % chunk_size
            
            with wave.open(self.file_path, 'rb') as wav_file:
                # Seek to the start frame
                wav_file.setpos(start_frame)
                
                # Initialize data array
                if self.sampwidth == 1:
                    data = np.zeros(frames_to_read, dtype=np.uint8)
                elif self.sampwidth == 2:
                    data = np.zeros(frames_to_read, dtype=np.int16)
                elif self.sampwidth == 4:
                    data = np.zeros(frames_to_read, dtype=np.int32)
                
                # Load in chunks
                for i in range(num_chunks):
                    raw_data = wav_file.readframes(chunk_size)
                    if not raw_data:
                        break
                    
                    # Copy data to appropriate position
                    if self.sampwidth == 1:
                        data[i*chunk_size:(i+1)*chunk_size] = np.frombuffer(raw_data, dtype=np.uint8)
                    elif self.sampwidth == 2:
                        data[i*chunk_size:(i+1)*chunk_size] = np.frombuffer(raw_data, dtype=np.int16)
                    elif self.sampwidth == 4:
                        data[i*chunk_size:(i+1)*chunk_size] = np.frombuffer(raw_data, dtype=np.int32)
                
                # Load remaining data
                if remaining > 0:
                    raw_data = wav_file.readframes(remaining)
                    if raw_data:
                        if self.sampwidth == 1:
                            data[-remaining:] = np.frombuffer(raw_data, dtype=np.uint8)
                        elif self.sampwidth == 2:
                            data[-remaining:] = np.frombuffer(raw_data, dtype=np.int16)
                        elif self.sampwidth == 4:
                            data[-remaining:] = np.frombuffer(raw_data, dtype=np.int32)
            
            # Convert to float32 and normalize
            if self.sampwidth == 1:
                data = data.astype(np.float32) / 128.0 - 1.0
            elif self.sampwidth == 2:
                data = data.astype(np.float32) / 32768.0
            elif self.sampwidth == 4:
                data = data.astype(np.float32) / (2**31)
            
            # Handle stereo
            if self.channels == 2:
                data = data[::2]  # Take left channel only
            
            # Store the current chunk
            self.current_chunk = data
            self.current_chunk_start = start_frame
            
            # Update processed data for the new chunk
            if self.abs_data is not None:
                self.abs_data = data.copy()
                self.abs_data = np.abs(self.abs_data)
            if self.smoothed_data is not None and self.abs_data is not None:
                window_size = 1000  # You might want to make this configurable
                self.smoothed_data = np.convolve(self.abs_data, 
                                                np.ones(window_size)/window_size, 
                                                mode='same')
            
            return data
            
        except KeyboardInterrupt:
            print("Chunk loading cancelled by user")
            return None
        except Exception as e:
            print(f"Error loading chunk: {e}")
            return None
    
    def update_plot(self):
        if not hasattr(self, 'file_path') or not self.file_path or not hasattr(self, 'wav_data'):
            return
        
        # Enable mouse interaction
        self.canvas.mpl_connect('button_press_event', self.on_mouse_press)
        self.canvas.mpl_connect('button_release_event', self.on_mouse_release)
        self.canvas.mpl_connect('motion_notify_event', self.on_mouse_move)
        
        # Clear the plot
        self.ax.clear()
        
        # Reset region lines references
        self.region_left_line = None
        self.region_right_line = None
        self.region_rect = None
        
        # Ensure view_start is within bounds
        self.view_start = max(0, min(self.view_start, self.total_frames - 1))
        
        # Calculate view range
        view_end = min(self.view_start + self.view_range, self.total_frames)
        
        # Get visible portion of data directly from the full dataset
        visible_data = self.wav_data[self.view_start:view_end]
    
        # Set y-axis limits based on data type
        if self.abs_data is not None:
            self.ax.set_ylim(0, 1)
        else:
            self.ax.set_ylim(-1, 1)
    
        # Plot the waveform
        if self.smoothed_data is not None:
            plot_data = self.smoothed_data[self.view_start:view_end]
            color, label = 'g-', 'Smoothed'
        elif self.abs_data is not None:
            plot_data = self.abs_data[self.view_start:view_end]
            color, label = 'b-', 'Absolute'
        else:
            plot_data = visible_data
            color, label = 'k-', 'Raw'
    
        # Calculate time values in milliseconds
        start_time_ms = self.view_start * 1000 / self.sample_rate
        end_time_ms = view_end * 1000 / self.sample_rate
    
        # Create time array for the plot data
        time_ms = np.linspace(start_time_ms, end_time_ms, len(plot_data))
    
        # Downsample if too many points
        max_points = 10000
        if len(plot_data) > max_points:
            downsample = len(plot_data) // max_points
            plot_data = plot_data[::downsample]
            time_ms = time_ms[::downsample]
        
        # Set x-axis limits to keep the view fixed
        self.ax.set_xlim(start_time_ms, end_time_ms)
    
        # Plot the data
        self.ax.plot(time_ms, plot_data, color, label=label, linewidth=0.5)
    
        # Always draw both threshold lines
        # Draw absolute threshold in red
        self.ax.axhline(y=self.abs_threshold, color='red', linestyle='-', label='Absolute Threshold')
    
        # Calculate and draw relative threshold in violet
        rel_threshold = self.rel_threshold
        if len(plot_data) > 0:
            rel_threshold = self.rel_threshold * np.max(np.abs(plot_data))
    
        self.ax.axhline(y=rel_threshold, color='violet', linestyle='-', label='Relative Threshold')
    
        # Set the active threshold based on current mode
        if self.using_absolute_threshold:
            self.threshold = self.abs_threshold
        else:
            self.threshold = rel_threshold
    
        # Update threshold labels
        self.threshold_label.setText(f"Threshold: {self.threshold:.3f}")
        self.threshold_mode_label.setText(f"Mode: {'Absolute' if self.using_absolute_threshold else 'Relative'}")
    
        # Plot detected pulses
        for pulse in self.pulses:
            pulse_pos = pulse['position']  # Get position from pulse dictionary
            if self.view_start <= pulse_pos < view_end:
                pulse_time = pulse_pos / self.sample_rate * 1000  # Convert to ms
                
                # Get pulse height directly from the data
                if self.smoothed_data is not None:
                    pulse_height = self.smoothed_data[pulse_pos]
                elif self.abs_data is not None:
                    pulse_height = self.abs_data[pulse_pos]
                else:
                    pulse_height = self.wav_data[pulse_pos]
                
                # Use different colors for positive and negative peaks
                peak_color = 'go' if pulse.get('peak_type') == 'negative' else 'ro'
                self.ax.plot(pulse_time, pulse_height, peak_color, markersize=5)
    
        # Plot detected skips
        for skip in self.skips:
            if self.view_start <= skip['position'] < view_end:  # Access position from skip dictionary
                skip_time = skip['position'] / self.sample_rate * 1000  # Convert to ms
                # Add magenta X marker
                self.ax.plot(skip_time, self.threshold, 'mx', markersize=10, markeredgewidth=2)
                # Add blue vertical line for skip
                self.ax.axvline(x=skip_time, color='blue', linestyle='--', alpha=0.7)
    
        self.ax.set_xlabel('Time (ms)')
        self.ax.set_ylabel('Amplitude')
        self.ax.set_title('Katydid Call Waveform Analysis')
        self.ax.grid(True)
        self.ax.legend(loc='upper right')  # Specify a fixed location for the legend
        # Redraw selection if active
        if self.selection_start is not None and self.selection_end is not None and self.selection_ystart is not None and self.selection_yend is not None:
            x_min = min(self.selection_start, self.selection_end)
            x_max = max(self.selection_start, self.selection_end)
            y_min = min(self.selection_ystart, self.selection_yend)
            y_max = max(self.selection_ystart, self.selection_yend)
        
            # Create rectangle for selection
            from matplotlib.patches import Rectangle
            if self.selection_rect:
                try:
                    self.selection_rect.remove()
                except:
                    pass  # In case the rectangle was already removed
        
            # Calculate normalized coordinates for the rectangle
            ymin, ymax = self.ax.get_ylim()
            y_range = ymax - ymin
        
            norm_y_min = (y_min - ymin) / y_range
            norm_y_max = (y_max - ymin) / y_range
        
            self.selection_rect = self.ax.axvspan(x_min, x_max, 
                                                ymin=norm_y_min, 
                                                ymax=norm_y_max,
                                                alpha=0.3, color='yellow')
        
        # Add on-graph time display in a small black box
        if self.selection_start is not None and self.selection_end is not None:
            duration = abs(self.selection_end - self.selection_start)
            self.time_text = self.ax.text(
                0.5, 0.03, 
                f"Selection: {duration:.3f} ms", 
                transform=self.ax.transAxes,
                ha='center', va='bottom',
                bbox=dict(
                    facecolor='black', 
                    alpha=0.7,
                    edgecolor='white',
                    boxstyle='round,pad=0.5'
                ),
                color='#00ff00',
                fontsize=9
            )
        
        # Draw region selection lines if active
        if hasattr(self, 'region_selection_active') and self.region_selection_active:
            if hasattr(self, 'region_left_pos') and self.region_left_pos is not None:
                # Draw left line
                self.region_left_line = self.ax.axvline(x=self.region_left_pos, color='red', linestyle='-', linewidth=2)
                
                # Draw right line if it's set
                if hasattr(self, 'region_right_pos') and self.region_right_pos is not None:
                    self.region_right_line = self.ax.axvline(x=self.region_right_pos, color='red', linestyle='-', linewidth=2)
                    
                    # Draw shaded region between lines
                    self.region_rect = self.ax.axvspan(self.region_left_pos, self.region_right_pos, alpha=0.2, color='green')
        
        self.canvas.draw()
    
    def on_mouse_press(self, event):
        if not hasattr(self, 'ax') or event.inaxes != self.ax:
            return
        
        # Set selection flag
        self.is_selecting = True
        
        # Store initial selection coordinates
        self.selection_start = event.xdata
        self.selection_end = event.xdata
        self.selection_ystart = event.ydata
        self.selection_yend = event.ydata
        
        # Remove existing selection rectangle if present
        if hasattr(self, 'selection_rect') and self.selection_rect:
            try:
                self.selection_rect.remove()
            except:
                pass
        
        # Draw a small starting rectangle
        if event.xdata is not None and event.ydata is not None:
            # Calculate normalized coordinates for the rectangle
            ymin, ymax = self.ax.get_ylim()
            y_range = ymax - ymin
            
            norm_y_min = (self.selection_ystart - ymin) / y_range
            norm_y_max = (self.selection_ystart - ymin) / y_range + 0.01  # Small height to start
            
            # Ensure values are within [0, 1] range
            norm_y_min = max(0, min(1, norm_y_min))
            norm_y_max = max(0, min(1, norm_y_max))
            
            self.selection_rect = self.ax.axvspan(event.xdata, event.xdata, 
                                                ymin=norm_y_min, 
                                                ymax=norm_y_max,
                                                alpha=0.3, color='yellow')
            
            # Update time display
            self.update_selection_time()
            
            # Use draw_idle() for more responsive updates
            self.canvas.draw_idle()
    
    def on_mouse_move(self, event):
        # Early exit if no selection is active or mouse outside the axes
        if not self.is_selecting:
            return
        
        if not hasattr(self, 'ax') or event.inaxes != self.ax:
            return
        
        # Update selection endpoint
        if event.xdata is not None:
            self.selection_end = event.xdata
        if event.ydata is not None:
            self.selection_yend = event.ydata
        
        # Remove existing rectangle if present
        if hasattr(self, 'selection_rect') and self.selection_rect:
            try:
                self.selection_rect.remove()
            except:
                pass
        
        # Check if selection_start or selection_end is None before comparing
        if self.selection_start is None or self.selection_end is None:
            return
        
        # Calculate normalized coordinates for the rectangle
        x_min = min(self.selection_start, self.selection_end)
        x_max = max(self.selection_start, self.selection_end)
        
        # Ensure y coordinates are valid
        if self.selection_ystart is None or self.selection_yend is None:
            y_min, y_max = 0, 1  # Use full height if y-coordinates aren't available
        else:
            y_min = min(self.selection_ystart, self.selection_yend)
            y_max = max(self.selection_ystart, self.selection_yend)
        
            ymin, ymax = self.ax.get_ylim()
            y_range = ymax - ymin
            
            norm_y_min = (y_min - ymin) / y_range
            norm_y_max = (y_max - ymin) / y_range
            
            # Ensure values are within [0, 1] range
            norm_y_min = max(0, min(1, norm_y_min))
            norm_y_max = max(0, min(1, norm_y_max))
        
        # Create the selection rectangle
        self.selection_rect = self.ax.axvspan(x_min, x_max,
                                            ymin=norm_y_min, 
                                            ymax=norm_y_max,
                                            alpha=0.3, color='yellow')
                                            
        # Update time text on graph
        if hasattr(self, 'time_text') and self.time_text is not None:
            try:
                self.time_text.remove()
            except:
                pass
        
        duration = abs(x_max - x_min)
        self.time_text = self.ax.text(
            0.5, 0.03, 
            f"Selection: {duration:.3f} ms", 
            transform=self.ax.transAxes,
            ha='center', va='bottom',
            bbox=dict(
                facecolor='black', 
                alpha=0.7,
                edgecolor='white',
                boxstyle='round,pad=0.5'
            ),
            color='#00ff00',
            fontsize=9
        )
        
        # Update time display
        self.update_selection_time()
        
        # We'll let update_plot handle drawing the region selection lines
        
        # Use draw_idle() for more efficient updates
        self.canvas.draw_idle()
    
    def on_mouse_release(self, event):
        # Only process if we were in a selecting state
        if not self.is_selecting:
            return
        
        try:
            # End selection
            self.is_selecting = False
            
            # Ensure we have a valid selection
            if self.selection_start is None or self.selection_end is None:
                return
            

            
            # Update final coordinates if mouse is within axes
            if event.inaxes == self.ax and event.xdata is not None:
                self.selection_end = event.xdata
                if event.ydata is not None:
                    self.selection_yend = event.ydata
            

            
            # Make sure we update the time indicator
            self.update_selection_time()
            
            # Draw on-graph time box if we have a valid selection
            x_min = min(self.selection_start, self.selection_end)
            x_max = max(self.selection_start, self.selection_end)
            duration = abs(x_max - x_min)
            
            if hasattr(self, 'time_text') and self.time_text is not None:
                try:
                    self.time_text.remove()
                except:
                    pass
            
            self.time_text = self.ax.text(
                0.5, 0.03, 
                f"Selection: {duration:.3f} ms", 
                transform=self.ax.transAxes,
                ha='center', va='bottom',
                bbox=dict(
                    facecolor='black', 
                    alpha=0.7,
                    edgecolor='white',
                    boxstyle='round,pad=0.5'
                ),
                color='#00ff00',
                fontsize=9
            )
            
            # Force complete redraw to ensure everything is visible
            self.canvas.draw()
        except Exception as e:
            print(f"Error in mouse release event: {e}")
    
    def add_manual_pulse(self):
        if self.selection_start is None or self.selection_end is None:
            return
        
        # Convert time to sample indices
        start_sample = int(min(self.selection_start, self.selection_end) * self.sample_rate / 1000)
        end_sample = int(max(self.selection_start, self.selection_end) * self.sample_rate / 1000)
        
        # Get amplitude bounds
        if self.selection_ystart is not None and self.selection_yend is not None:
            min_amp = min(self.selection_ystart, self.selection_yend)
            max_amp = max(self.selection_ystart, self.selection_yend)
        else:
            min_amp = float('-inf')
            max_amp = float('inf')
        
        # Get the data in the selection
        if self.smoothed_data is not None:
            data = self.smoothed_data[start_sample:end_sample]
        elif self.abs_data is not None:
            data = self.abs_data[start_sample:end_sample]
        else:
            data = self.wav_data[start_sample:end_sample]
        
        # Find multiple peaks in the selection that are within amplitude bounds
        if len(data) > 0:
            # Filter data by amplitude bounds
            valid_indices = []
            for i, amplitude in enumerate(data):
                if min_amp <= amplitude <= max_amp:
                    valid_indices.append(i)
            
            if valid_indices:
                # Find local peaks among valid indices
                # Define what constitutes a local peak - approximately 1ms of samples
                peak_width = int(0.5 * self.sample_rate / 1000)  # 0.5ms in samples
                if peak_width < 1:
                    peak_width = 1
                
                # Find all local peaks
                peaks = []
                for idx in valid_indices:
                    is_peak = True
                    # Check if this is a local maximum within peak_width
                    for offset in range(1, peak_width + 1):
                        # Check if any sample within peak_width is higher
                        if idx - offset >= 0 and idx - offset < len(data) and data[idx] < data[idx - offset]:
                            is_peak = False
                            break
                        if idx + offset < len(data) and data[idx] < data[idx + offset]:
                            is_peak = False
                            break
                
                    if is_peak:
                        peaks.append(idx)
            
                # Add all found peaks to pulses
                added_count = 0
                for peak_idx in peaks:
                    global_peak_idx = start_sample + peak_idx
                    # Check if a pulse already exists close to this location
                    duplicate = False
                    for pulse in self.pulses:
                        if abs(pulse['position'] - global_peak_idx) < peak_width:
                            duplicate = True
                            break
                
                    if not duplicate:
                        self.pulses.append({
                            'position': global_peak_idx,
                            'type': 'manual'
                        })
                        added_count += 1
            
                # Sort pulses
                self.pulses.sort(key=lambda x: x['position'])
                

                
                # Clear selection
                self.selection_start = None
                self.selection_end = None
                self.selection_ystart = None
                self.selection_yend = None
                if self.selection_rect:
                    try:
                        self.selection_rect.remove()
                        self.selection_rect = None
                    except:
                        pass
                self.update_plot()
                
                if added_count > 0:
                    QMessageBox.information(self, "Added Pulses", f"Added {added_count} new pulse(s) to the analysis.")
                else:
                    QMessageBox.warning(self, "No New Pulses", "No new pulses were added. Any detected peaks may already exist in the analysis.")
            else:
                QMessageBox.warning(self, "No Valid Amplitudes", "No valid amplitude values found within the selected range.")
    
    def delete_selected_pulses(self):
        """Delete pulses within the current selection."""
        if self.selection_start is None or self.selection_end is None:
            return
        
        # Convert time to sample indices
        start_sample = int(min(self.selection_start, self.selection_end) * self.sample_rate / 1000)
        end_sample = int(max(self.selection_start, self.selection_end) * self.sample_rate / 1000)
        
        # Get amplitude bounds
        if self.selection_ystart is not None and self.selection_yend is not None:
            min_amp = min(self.selection_ystart, self.selection_yend)
            max_amp = max(self.selection_ystart, self.selection_yend)
        else:
            min_amp = float('-inf')
            max_amp = float('inf')
            

            
        # Find pulses within the selection time range
        pulses_in_range = []
        for p in self.pulses:
            if start_sample <= p['position'] <= end_sample:
                # Check if pulse amplitude is within bounds
                if self.smoothed_data is not None:
                    amp = self.smoothed_data[p['position']]
                elif self.abs_data is not None:
                    amp = self.abs_data[p['position']]
                else:
                    amp = self.wav_data[p['position']]
                
                if min_amp <= amp <= max_amp:
                    pulses_in_range.append(p)
        
        # Remove pulses in range
        pulses_to_keep = [p for p in self.pulses if p not in pulses_in_range]
        pulses_removed = len(self.pulses) - len(pulses_to_keep)
        
        if pulses_removed > 0:
            self.pulses = pulses_to_keep
            QMessageBox.information(self, "Pulses Deleted", f"Removed {pulses_removed} pulse(s) from selection.")
        else:
            QMessageBox.information(self, "No Pulses Found", "No pulses were found in the selected area.")
        
        # Clear selection
        self.selection_start = None
        self.selection_end = None
        self.selection_ystart = None
        self.selection_yend = None
        if self.selection_rect:
            try:
                self.selection_rect.remove()
            except:
                pass
            self.selection_rect = None
        
        self.update_plot()
    
    def move_view(self, direction):
        """Move the view by a fixed amount in the specified direction (-1 for left, 1 for right)"""
        # Move by 25% of the view range
        move_amount = self.view_range // 4
        new_start = self.view_start + (direction * move_amount)
        
        # Ensure we stay within bounds
        new_start = max(0, min(new_start, self.total_frames - self.view_range))
        
        self.view_start = new_start
        self.update_plot()
    
    def zoom_view(self, factor, center=None):
        if not hasattr(self, 'file_path') or not self.file_path:
            return
        
        # If no center point provided, use the middle of current view
        if center is None:
            center = self.view_start + self.view_range // 2
        
        if factor < 1:  # Zooming in
            new_range = max(100, int(self.view_range * 0.5))  # Always halve the view when zooming in
        else:  # Zooming out
            new_range = min(int(self.view_range * 2), self.total_frames)  # Always double when zooming out
        
        # Use fixed zoom factors for consistent scaling
        # Calculate center point of current view
        # center = self.view_start + self.view_range // 2
        
        # Adjust view_start to maintain center point
        new_start = center - new_range // 2
        new_start = max(0, min(new_start, self.total_frames - new_range))
        
        self.view_range = new_range
        self.view_start = new_start
        self.update_plot()
        
    def pan_view(self, direction, amount=0.5):
        """Pan the view left or right by a percentage of the current view range.
        
        Args:
            direction: 'left' or 'right'
            amount: Amount to pan as a fraction of the current view (0.0-1.0)
        """
        if not hasattr(self, 'file_path') or not self.file_path:
            return
            
        # Calculate the pan distance in samples
        pan_distance = int(self.view_range * amount)
        
        if direction == 'left':
            # Pan left (decrease view_start)
            self.view_start = max(0, self.view_start - pan_distance)
        elif direction == 'right':
            # Pan right (increase view_start)
            max_start = max(0, self.total_frames - self.view_range)
            self.view_start = min(max_start, self.view_start + pan_distance)
        
        self.update_plot()

    def invert_values(self):
        """Invert the waveform by flipping positive and negative values."""
        if not hasattr(self, 'file_path') or not self.file_path:
            return
        
        # Invert the waveform (multiply by -1)
        self.inversion_count += 1
        
        # Directly invert the original waveform data
        self.wav_data = -1 * self.wav_data
        
        # Reset all processing data
        self.abs_data = None
        self.smoothed_data = None
        
        # Clear all detected pulses and skips
        self.pulses = []
        self.skips = []
        

        
        # Update the plot
        self.update_plot()
        
    def keyPressEvent(self, event):
        """Handle keyboard events."""
        from PyQt5.QtCore import Qt
        
        # Only process if a file is loaded
        if not hasattr(self, 'file_path') or not self.file_path:
            super().keyPressEvent(event)  # Pass to parent for default handling
            return
            
        key = event.key()
        
        # WASD navigation
        if key == Qt.Key_W:
            # Zoom in
            self.zoom_view(0.5)
        elif key == Qt.Key_S:
            # Zoom out
            self.zoom_view(2.0)
        elif key == Qt.Key_A:
            # Move left
            self.move_view(-1)
        elif key == Qt.Key_D:
            # Move right
            self.move_view(1)
        
        # Region selection with / key
        elif key == Qt.Key_Slash:
            if not self.region_selection_active:
                # Start region selection mode
                self.region_selection_active = True
                self.region_selection_mode = 'left'
                
                # Create initial left line at 1/4 of the view
                view_time_start = self.view_start / self.sample_rate * 1000
                view_time_end = (self.view_start + self.view_range) / self.sample_rate * 1000
                view_width = view_time_end - view_time_start
                
                self.region_left_pos = view_time_start + view_width * 0.25
                self.region_right_pos = view_time_start + view_width * 0.75
                
                # Update the plot to show the region lines
                self.update_plot()
                self.show_status_message("Use left/right arrow keys to position the left line, press Enter to set")
            else:
                # Exit region selection mode
                self.region_selection_active = False
                self.region_selection_mode = None
                self.region_left_pos = None
                self.region_right_pos = None
                self.update_plot()
                self.show_status_message("Region selection canceled")
        
        # Arrow keys for region line positioning
        elif self.region_selection_active and (key == Qt.Key_Left or key == Qt.Key_Right):
            # Calculate movement amount (1% of view width)
            view_time_start = self.view_start / self.sample_rate * 1000
            view_time_end = (self.view_start + self.view_range) / self.sample_rate * 1000
            view_width = view_time_end - view_time_start
            move_amount = view_width * 0.01
            
            if self.region_selection_mode == 'left':
                # Move left line
                if key == Qt.Key_Left:
                    self.region_left_pos = max(view_time_start, self.region_left_pos - move_amount)
                else:  # Right key
                    if self.region_right_pos is not None:
                        # Don't move past right line
                        self.region_left_pos = min(self.region_right_pos - move_amount, self.region_left_pos + move_amount)
                    else:
                        self.region_left_pos = min(view_time_end, self.region_left_pos + move_amount)
            else:  # 'right' mode
                # Move right line
                if key == Qt.Key_Left:
                    # Don't move past left line
                    self.region_right_pos = max(self.region_left_pos + move_amount, self.region_right_pos - move_amount)
                else:  # Right key
                    self.region_right_pos = min(view_time_end, self.region_right_pos + move_amount)
            
            # Update the plot
            self.update_plot()
        
        # Enter key to confirm region line position
        elif self.region_selection_active and key == Qt.Key_Return:
            if self.region_selection_mode == 'left':
                # Left line set, now set right line
                self.region_selection_mode = 'right'
                self.show_status_message("Use left/right arrow keys to position the right line, press Enter to set")
            elif self.region_selection_mode == 'right':
                # Right line set, now set to 'complete' mode for the third Enter press
                self.region_selection_mode = 'complete'
                self.show_status_message("Press Enter again to detect pulses in this region and exit selection mode")
            else:  # 'complete' mode - third Enter press
                # Detect pulses in the selected region
                self.detect_pulses()
                
                # Exit region selection mode
                self.region_selection_active = False
                self.region_selection_mode = None
                self.show_status_message("Pulses detected in selected region. Region selection mode exited.")
            
            # Update the plot
            self.update_plot()
            
        # Other keyboard shortcuts
        elif key == Qt.Key_Y:
            # Detect pulses
            self.detect_pulses()
        elif key == Qt.Key_T:
            # Analyze periods
            self.analyze_pulse_periods()
        elif key == Qt.Key_R:
            # Invert values
            self.invert_values()
        elif key == Qt.Key_G:
            # Smooth signal
            self.apply_smoothing()

        elif key == Qt.Key_O:
            # Add pulse
            self.add_manual_pulse()
        elif key == Qt.Key_P:
            # Delete pulse
            self.delete_selected_pulses()
        elif key == Qt.Key_Up:
            # Increase threshold
            if self.using_absolute_threshold:
                self.abs_threshold = min(1.0, self.abs_threshold + 0.025)
            else:
                self.rel_threshold = min(1.0, self.rel_threshold + 0.025)
            self.update_plot()
            self.show_status_message(f"Threshold increased to {self.abs_threshold if self.using_absolute_threshold else self.rel_threshold:.3f}")
        elif key == Qt.Key_Down:
            # Decrease threshold
            if self.using_absolute_threshold:
                self.abs_threshold = max(0.0, self.abs_threshold - 0.025)
            else:
                self.rel_threshold = max(0.0, self.rel_threshold - 0.025)
            self.update_plot()
            self.show_status_message(f"Threshold decreased to {self.abs_threshold if self.using_absolute_threshold else self.rel_threshold:.3f}")
        elif key == Qt.Key_Equal:
            # Save results with WAV file
            self.save_results_with_wav()
        elif key == Qt.Key_BracketLeft:
            # Switch to relative threshold mode
            self.using_absolute_threshold = False
            self.threshold_mode_label.setText("Mode: Relative")
            self.threshold = self.rel_threshold
            self.update_plot()
            self.show_status_message("Switched to relative threshold mode")
        elif key == Qt.Key_BracketRight:
            # Switch to absolute threshold mode
            self.using_absolute_threshold = True
            self.threshold_mode_label.setText("Mode: Absolute")
            self.threshold = self.abs_threshold
            self.update_plot()
            self.show_status_message("Switched to absolute threshold mode")
        elif key == Qt.Key_F11:
            if self.isFullScreen():
                self.showNormal()
            else:
                self.showFullScreen()
        elif key == Qt.Key_Escape and self.isFullScreen():
            self.showNormal()
        else:
            super().keyPressEvent(event)
    
    def apply_smoothing(self):
        """Apply smoothing to the waveform data."""
        if not hasattr(self, 'file_path') or not self.file_path:
            return
            
        # Create a copy of the current waveform data for smoothing
        if self.abs_data is None:
            # If no processed data exists, create it from the current waveform
            self.abs_data = self.wav_data.copy()
        
        # Calculate window size (much smaller - 0.25ms window, 1/20 of original)
        # Adjust this value based on your system's memory capacity
        window_size = int(self.sample_rate * 0.00025)  # 0.25ms instead of 5ms
        if window_size % 2 == 0:
            window_size += 1  # Make sure window size is odd
        window_size = max(3, window_size)  # Ensure minimum size of 3
        
        # If smoothed_data already exists, smooth it further
        # Otherwise, start with abs_data
        source_data = self.smoothed_data if self.smoothed_data is not None else self.abs_data
        
        # Create window and apply convolution
        window = np.ones(window_size) / window_size
        self.smoothed_data = np.zeros_like(source_data)
        
        # Process in segments to avoid memory issues
        segment_size = min(len(source_data), 1000000)  # Process 1M samples at a time
        for i in range(0, len(source_data), segment_size):
            end = min(i + segment_size, len(source_data))
            self.smoothed_data[i:end] = np.convolve(
                source_data[i:end],
                window,
                mode='same'
            )
        
        # Don't reset pulses when smoothing multiple times
        self.update_plot()

    def detect_pulses(self):
        """Detect pulses in the entire waveform."""
        # Initialize pulse detection
        new_pulses = []
        
        # Determine which data to use for pulse detection - ALWAYS use raw data
        detection_data = self.wav_data.copy()  # Use a copy to avoid modifying original data
            
        # Get current threshold - ensure we're using the correct threshold mode
        if self.using_absolute_threshold:
            threshold = self.abs_threshold
        else:
            # For relative threshold, calculate based on max value in entire waveform
            if len(detection_data) > 0:
                threshold = self.rel_threshold * np.max(detection_data)
            else:
                threshold = self.rel_threshold
                
        # Determine if we're looking for negative or positive peaks based on threshold sign
        looking_for_negative_peaks = threshold < 0
                
        # Print threshold for debugging
        print(f"Threshold: {threshold}, Looking for {'NEGATIVE' if looking_for_negative_peaks else 'POSITIVE'} peaks")
        
        # Define detection range - use entire waveform
        start_idx = 0
        end_idx = len(detection_data)
        
        # If region selection is active and both lines are set, limit detection to that region
        if hasattr(self, 'region_selection_active') and self.region_selection_active and \
           hasattr(self, 'region_left_pos') and self.region_left_pos is not None and \
           hasattr(self, 'region_right_pos') and self.region_right_pos is not None:
            # Convert time positions (ms) to sample indices
            region_start_sample = int(self.region_left_pos * self.sample_rate / 1000)
            region_end_sample = int(self.region_right_pos * self.sample_rate / 1000)
            
            # Make sure they're within valid range
            region_start_sample = max(0, region_start_sample)
            region_end_sample = min(len(detection_data), region_end_sample)
            
            # Update detection range
            start_idx = region_start_sample
            end_idx = region_end_sample
        
        # Find peaks based on threshold sign
        peaks = []
        current_peak = None
        in_peak = False
        
        # First pass: find peaks based on threshold sign
        for i in range(start_idx, end_idx):
            if looking_for_negative_peaks:
                # Looking for NEGATIVE peaks below threshold (threshold is negative)
                if detection_data[i] < 0 and detection_data[i] < threshold:
                    # We're in a potential peak region
                    in_peak = True
                    
                    # If this is the first point below threshold or lower than current peak
                    if current_peak is None or detection_data[i] < detection_data[current_peak]:
                        current_peak = i
                else:
                    # We're above threshold or positive
                    if in_peak and current_peak is not None:
                        # We just left a peak region, add the peak we found
                        # Triple check that the peak is negative and below threshold
                        if detection_data[current_peak] < 0 and detection_data[current_peak] < threshold:
                            peaks.append(current_peak)
                            print(f"Found NEGATIVE peak at {current_peak} with value {detection_data[current_peak]} < threshold {threshold}")
                        current_peak = None
                    in_peak = False
            else:
                # Looking for POSITIVE peaks above threshold (threshold is positive)
                if detection_data[i] > 0 and detection_data[i] > threshold:
                    # We're in a potential peak region
                    in_peak = True
                    
                    # If this is the first point above threshold or higher than current peak
                    if current_peak is None or detection_data[i] > detection_data[current_peak]:
                        current_peak = i
                else:
                    # We're below threshold or negative
                    if in_peak and current_peak is not None:
                        # We just left a peak region, add the peak we found
                        # Triple check that the peak is positive and above threshold
                        if detection_data[current_peak] > 0 and detection_data[current_peak] > threshold:
                            peaks.append(current_peak)
                            print(f"Found POSITIVE peak at {current_peak} with value {detection_data[current_peak]} > threshold {threshold}")
                        current_peak = None
                    in_peak = False
        
        # Add last peak if we ended while still in a peak
        if in_peak and current_peak is not None:
            if looking_for_negative_peaks:
                # Triple check that the peak is negative and below threshold
                if detection_data[current_peak] < 0 and detection_data[current_peak] < threshold:
                    peaks.append(current_peak)
            else:
                # Triple check that the peak is positive and above threshold
                if detection_data[current_peak] > 0 and detection_data[current_peak] > threshold:
                    peaks.append(current_peak)
        
        # Second pass: filter out close peaks
        min_distance = int(self.sample_rate * 0.001)  # Minimum 1ms apart
        filtered_peaks = []
        
        for peak in peaks:
            # Check if this peak is too close to the previous one
            if not filtered_peaks or (peak - filtered_peaks[-1]) >= min_distance:
                filtered_peaks.append(peak)
        
        # Add the new pulses
        for peak in filtered_peaks:
            new_pulses.append({
                'position': peak,
                'type': 'detected',
                'peak_type': 'negative' if looking_for_negative_peaks else 'positive'
            })
        

        
        # Update pulses
        self.pulses.extend(new_pulses)
        self.update_plot()
    
    def analyze_pulse_periods(self):
        """
        Analyze pulse periods to find patterns in the data.
        For each period (group of 3 pulses), calculate:
        1. Total duration between pulse 1 and 3
        2. Ratio of time between pulses 1 and 2 divided by the total period duration
        3. Amplitude and time information for each pulse
        """
        if not self.pulses or len(self.pulses) < 3:
            QMessageBox.warning(self, "Period Analysis", "Not enough pulses detected. Please detect pulses first.")
            return
            
        # Sort pulses by position to ensure proper ordering
        sorted_pulses = sorted(self.pulses, key=lambda p: p['position'])
        
        # Calculate periods - each period contains 3 pulses (1-3, 2-4, etc.)
        periods = []
        individual_pulses = []
        
        # First, process all individual pulses
        for i, pulse in enumerate(sorted_pulses):
            # Get amplitude values for the pulse
            pulse_amp = self.wav_data[pulse['position']] if pulse['position'] < len(self.wav_data) else 0
            
            # Calculate time in milliseconds for the pulse
            pulse_time = pulse['position'] / self.sample_rate * 1000  # ms
            
            individual_pulses.append({
                'index': i+1,
                'time': pulse_time,
                'amplitude': pulse_amp,
                'position': pulse['position']
            })
        
        # Then, calculate periods using groups of 3 pulses
        for i in range(len(sorted_pulses) - 2):  # Need at least 3 pulses for a period
            pulse1 = sorted_pulses[i]
            pulse2 = sorted_pulses[i+1]
            pulse3 = sorted_pulses[i+2]
            
            # Calculate period duration and pulse ratio
            period_duration = (pulse3['position'] - pulse1['position']) / self.sample_rate * 1000  # ms
            pulse_interval = (pulse2['position'] - pulse1['position']) / self.sample_rate * 1000  # ms
            pulse_ratio = pulse_interval / period_duration if period_duration > 0 else 0
            
            periods.append({
                'index': i+1,
                'duration': period_duration,
                'ratio': pulse_ratio
            })
            
        # Store the periods and pulses for later saving
        self.current_periods = periods
        self.current_pulses = individual_pulses
            
        # Create and show the analysis window
        self._show_period_analysis(periods, individual_pulses)
    
    def _show_period_analysis(self, periods, individual_pulses):
        """Display the period analysis in a new window with table and histograms"""
        # Create the dialog window
        analysis_window = QDialog(self)
        analysis_window.setWindowTitle("Pulse Period Analysis")
        analysis_window.resize(800, 600)
        
        # Create layout
        layout = QVBoxLayout(analysis_window)
        
        # Add tabs for different views
        tabs = QTabWidget()
        
        # Tab 1: Table view with 5 columns: Period, Duration, Pulse Ratio, Amplitude, Time
        table_tab = QWidget()
        table_layout = QVBoxLayout(table_tab)
        
        # Create and populate the table
        table = QTableWidget()
        table.setColumnCount(5)
        table.setHorizontalHeaderLabels(["Period", "Duration (ms)", "Pulse Ratio", "Amplitude", "Time (ms)"])
        
        # Calculate statistics for highlighting outliers
        durations = [p['duration'] for p in periods]
        
        # Calculate mode of durations for highlighting outliers
        if durations:
            hist, bin_edges = np.histogram(durations, bins=30)
            mode_bin_index = np.argmax(hist)
            mode_range = (bin_edges[mode_bin_index], bin_edges[mode_bin_index + 1])
            mode_duration = (mode_range[0] + mode_range[1]) / 2
        else:
            mode_duration = 0
        
        std_duration = np.std(durations) if durations else 0
        
        # Add data to the table - we'll have more rows than periods since each pulse gets its own row
        total_rows = len(individual_pulses)
        table.setRowCount(total_rows)
        
        # Populate the table with individual pulse data
        for row, pulse in enumerate(individual_pulses):
            # Find if this pulse is part of a period
            period_index = ""
            duration = ""
            ratio = ""
            
            # Check if this pulse is the first pulse in any period
            for p in periods:
                if p['index'] == pulse['index']:
                    period_index = str(p['index'])
                    duration = f"{p['duration']:.2f}"
                    ratio = f"{p['ratio']:.4f}"
                    break
            
            # Set the data
            table.setItem(row, 0, QTableWidgetItem(period_index))
            table.setItem(row, 1, QTableWidgetItem(duration))
            table.setItem(row, 2, QTableWidgetItem(ratio))
            table.setItem(row, 3, QTableWidgetItem(f"{pulse['amplitude']:.4f}"))
            table.setItem(row, 4, QTableWidgetItem(f"{pulse['time']:.2f}"))
            
            # Highlight outliers (more than 2 standard deviations from MODE) if this is a period
            if period_index and duration:
                deviation = float(duration) - mode_duration
                if abs(deviation) > 2 * std_duration and std_duration > 0:
                    for col in range(5):
                        item = table.item(row, col)
                        item.setBackground(QColor(255, 200, 200))  # Light red background
        
        # Auto-adjust column widths
        table.resizeColumnsToContents()
        
        # Add table to layout
        table_layout.addWidget(table)
        
        # Tab 2: Histogram of period durations showing MODE
        duration_tab = QWidget()
        duration_layout = QVBoxLayout(duration_tab)
        
        duration_figure = Figure(figsize=(5, 4), tight_layout=True)
        duration_canvas = FigureCanvas(duration_figure)
        duration_ax = duration_figure.add_subplot(111)
        
        # Extract duration data
        durations = [p['duration'] for p in periods]
        
        # Calculate mode of durations
        if durations:
            # Create bins for histogram
            hist, bin_edges = np.histogram(durations, bins=30)
            # Find the bin with the highest count
            mode_bin_index = np.argmax(hist)
            # Get the mode range
            mode_range = (bin_edges[mode_bin_index], bin_edges[mode_bin_index + 1])
            mode_value = (mode_range[0] + mode_range[1]) / 2
            
            # Plot histogram of durations
            n, bins, patches = duration_ax.hist(durations, bins=30, alpha=0.7, color='green')
            
            # Highlight the mode bin
            for i, patch in enumerate(patches):
                if i == mode_bin_index:
                    patch.set_facecolor('red')  # Highlight the mode bin
            
            # Add a vertical line at the mode
            duration_ax.axvline(x=mode_value, color='red', linestyle='--', linewidth=2)
            duration_ax.text(mode_value, max(n)*0.9, f'Mode: {mode_value:.2f} ms', 
                         color='red', fontweight='bold', ha='right')
        else:
            duration_ax.text(0.5, 0.5, 'No data available', ha='center', va='center', transform=duration_ax.transAxes)
        
        duration_ax.set_xlabel('Period Duration (ms)')
        duration_ax.set_ylabel('Frequency')
        duration_ax.set_title('Distribution of Period Durations (Mode Highlighted)')
        duration_ax.grid(True)
        duration_ax.legend()
        duration_ax.legend(loc='upper right')  # Specify a fixed location for the legend
        duration_figure.tight_layout()
        duration_canvas.draw()
        
        duration_layout.addWidget(duration_canvas)
        
        # Tab 3: Histogram of pulse ratios showing MODE
        ratio_tab = QWidget()
        ratio_layout = QVBoxLayout(ratio_tab)
        
        ratio_figure = Figure(figsize=(5, 4), tight_layout=True)
        ratio_canvas = FigureCanvas(ratio_figure)
        ratio_ax = ratio_figure.add_subplot(111)
        
        # Extract ratio data
        ratios = [p['ratio'] for p in periods]
        
        # Calculate mode of ratios
        if ratios:
            # Create bins for histogram
            hist, bin_edges = np.histogram(ratios, bins=30)
            # Find the bin with the highest count
            mode_bin_index = np.argmax(hist)
            # Get the mode range
            mode_range = (bin_edges[mode_bin_index], bin_edges[mode_bin_index + 1])
            mode_value = (mode_range[0] + mode_range[1]) / 2
            
            # Plot histogram of ratios
            n, bins, patches = ratio_ax.hist(ratios, bins=30, alpha=0.7, color='blue')
            
            # Highlight the mode bin
            for i, patch in enumerate(patches):
                if i == mode_bin_index:
                    patch.set_facecolor('red')  # Highlight the mode bin
            
            # Add a vertical line at the mode
            ratio_ax.axvline(x=mode_value, color='red', linestyle='--', linewidth=2)
            ratio_ax.text(mode_value, max(n)*0.9, f'Mode: {mode_value:.4f}', 
                       color='red', fontweight='bold', ha='right')
        else:
            ratio_ax.text(0.5, 0.5, 'No data available', ha='center', va='center', transform=ratio_ax.transAxes)
        
        ratio_ax.set_xlabel('Pulse Ratio (time between pulses 1-2 / period duration)')
        ratio_ax.set_ylabel('Frequency')
        ratio_ax.set_title('Distribution of Pulse Ratios (Mode Highlighted)')
        ratio_ax.grid(True)
        ratio_ax.legend()
        ratio_ax.legend(loc='upper right')  # Specify a fixed location for the legend
        ratio_figure.tight_layout()
        ratio_canvas.draw()
        
        ratio_layout.addWidget(ratio_canvas)
        
        # Tab 4: Statistics
        stats_tab = QWidget()
        stats_layout = QVBoxLayout(stats_tab)
        
        # Create a text browser for statistics
        stats_text = QTextBrowser()
        
        # Calculate statistics
        durations = [p['duration'] for p in periods]
        ratios = [p['ratio'] for p in periods]
        
        # Format statistics text
        stats_html = "<h2>Period Statistics</h2>"
        stats_html += "<h3>Duration Statistics (ms)</h3>"
        stats_html += f"<p>Count: {len(durations)}</p>"
        stats_html += f"<p>Mean: {np.mean(durations):.2f}</p>"
        stats_html += f"<p>Median: {np.median(durations):.2f}</p>"
        stats_html += f"<p>Mode: {mode_value if 'mode_value' in locals() else 0:.2f}</p>"
        stats_html += f"<p>Std Dev: {np.std(durations):.2f}</p>"
        stats_html += f"<p>Min: {np.min(durations):.2f}</p>"
        stats_html += f"<p>Max: {np.max(durations):.2f}</p>"
        
        stats_html += "<h3>Pulse Ratio Statistics</h3>"
        stats_html += f"<p>Mean: {np.mean(ratios):.4f}</p>"
        stats_html += f"<p>Median: {np.median(ratios):.4f}</p>"
        stats_html += f"<p>Mode: {mode_value if 'mode_value' in locals() else 0:.4f}</p>"
        stats_html += f"<p>Std Dev: {np.std(ratios):.4f}</p>"
        stats_html += f"<p>Min: {np.min(ratios):.4f}</p>"
        stats_html += f"<p>Max: {np.max(ratios):.4f}</p>"
        
        stats_text.setHtml(stats_html)
        stats_layout.addWidget(stats_text)
        
        # Add tabs to the tab widget
        tabs.addTab(table_tab, "Pulse Table")
        tabs.addTab(duration_tab, "Period Histogram")
        tabs.addTab(ratio_tab, "Ratio Histogram")
        tabs.addTab(stats_tab, "Statistics")
        
        # Add tab widget to main layout
        layout.addWidget(tabs)
        
        # Add close button
        close_button = QPushButton("Close")
        close_button.clicked.connect(analysis_window.close)
        layout.addWidget(close_button)
        
        # Show the window
        analysis_window.setLayout(layout)
        analysis_window.show()
        
    def save_results_with_wav(self):
        """Save analysis results to a folder with user-specified name, including WAV file."""
        # Check if we have pulses to save
        if not hasattr(self, 'current_pulses') or not self.current_pulses:
            QMessageBox.warning(self, "No Data to Save", "Please analyze pulse periods first (press T) before saving.")
            return
            
        # Ask user for a folder name
        folder_name, ok = QInputDialog.getText(self, "Save Results", "What do you want your folder name to be?", 
                                           QLineEdit.Normal, "")
        if not ok or not folder_name:
            return
            
        # Ask user for a directory to save the folder in
        save_dir = QFileDialog.getExistingDirectory(self, "Select Directory to Save Folder")
        if not save_dir:  # User canceled
            return
            
        # Create the folder
        folder_path = os.path.join(save_dir, folder_name)
        try:
            os.makedirs(folder_path, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "Error Creating Folder", f"Failed to create folder: {str(e)}")
            return
            
        try:
            # Create timestamp for unique filenames if needed
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # 1. Save the table as CSV with the requested format (5 columns)
            csv_file = os.path.join(folder_path, f"{folder_name}_table.csv")
            with open(csv_file, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Period', 'Duration (ms)', 'Pulse Ratio', 'Amplitude', 'Time (ms)'])
                
                # Write all individual pulse data
                for pulse in self.current_pulses:
                    # Find if this pulse is part of a period
                    period_index = ""
                    duration = ""
                    ratio = ""
                    
                    # Check if this pulse is the first pulse in any period
                    for p in self.current_periods:
                        if p['index'] == pulse['index']:
                            period_index = str(p['index'])
                            duration = f"{p['duration']:.2f}"
                            ratio = f"{p['ratio']:.4f}"
                            break
                    
                    # Write the row
                    writer.writerow([
                        period_index,
                        duration,
                        ratio,
                        f"{pulse['amplitude']:.4f}",
                        f"{pulse['time']:.2f}"
                    ])
            
            # 2. Save the period duration histogram
            period_hist_file = os.path.join(folder_path, f"{folder_name}_period_histogram.png")
            
            # Create figure for period histogram
            period_fig = Figure(figsize=(8, 6))
            period_canvas = FigureCanvas(period_fig)
            period_ax = period_fig.add_subplot(111)
            
            # Extract duration data
            durations = [p['duration'] for p in self.current_periods]
            
            # Calculate mode of durations
            if durations:
                # Create bins for histogram
                hist, bin_edges = np.histogram(durations, bins=30)
                # Find the bin with the highest count
                mode_bin_index = np.argmax(hist)
                # Get the mode range
                mode_range = (bin_edges[mode_bin_index], bin_edges[mode_bin_index + 1])
                mode_value = (mode_range[0] + mode_range[1]) / 2
                
                # Plot histogram of durations
                n, bins, patches = period_ax.hist(durations, bins=30, alpha=0.7, color='green')
                
                # Highlight the mode bin
                for i, patch in enumerate(patches):
                    if i == mode_bin_index:
                        patch.set_facecolor('red')  # Highlight the mode bin
            
                # Add a vertical line at the mode
                period_ax.axvline(x=mode_value, color='red', linestyle='--', linewidth=2)
                period_ax.text(mode_value, max(n)*0.9, f'Mode: {mode_value:.2f} ms', 
                            color='red', fontweight='bold', ha='right')
            else:
                period_ax.text(0.5, 0.5, 'No data available', ha='center', va='center', transform=period_ax.transAxes)
            
            period_ax.set_xlabel('Period Duration (ms)')
            period_ax.set_ylabel('Frequency')
            period_ax.set_title('Distribution of Period Durations (Mode Highlighted)')
            period_ax.grid(True)
            period_ax.legend(loc='upper right')
            period_fig.tight_layout()
            period_fig.savefig(period_hist_file)
            
            # 3. Save the ratio histogram
            ratio_hist_file = os.path.join(folder_path, f"{folder_name}_ratio_histogram.png")
            ratio_fig = Figure(figsize=(8, 6))
            ratio_canvas = FigureCanvas(ratio_fig)
            ratio_ax = ratio_fig.add_subplot(111)
            
            # Extract ratio data
            ratios = [p['ratio'] for p in self.current_periods]
            
            # Calculate mode of ratios
            if ratios:
                # Create bins for histogram
                hist, bin_edges = np.histogram(ratios, bins=30)
                # Find the bin with the highest count
                mode_bin_index = np.argmax(hist)
                # Get the mode range
                mode_range = (bin_edges[mode_bin_index], bin_edges[mode_bin_index + 1])
                mode_value = (mode_range[0] + mode_range[1]) / 2
                
                # Plot histogram of ratios
                n, bins, patches = ratio_ax.hist(ratios, bins=30, alpha=0.7, color='blue')
                
                # Highlight the mode bin
                for i, patch in enumerate(patches):
                    if i == mode_bin_index:
                        patch.set_facecolor('red')  # Highlight the mode bin
            
                # Add a vertical line at the mode
                ratio_ax.axvline(x=mode_value, color='red', linestyle='--', linewidth=2)
                ratio_ax.text(mode_value, max(n)*0.9, f'Mode: {mode_value:.4f}', 
                           color='red', fontweight='bold', ha='right')
            else:
                ratio_ax.text(0.5, 0.5, 'No data available', ha='center', va='center', transform=ratio_ax.transAxes)
            
            ratio_ax.set_xlabel('Pulse Ratio (time between pulses 1-2 / period duration)')
            ratio_ax.set_ylabel('Frequency')
            ratio_ax.set_title('Distribution of Pulse Ratios (Mode Highlighted)')
            ratio_ax.grid(True)
            ratio_ax.legend(loc='upper right')
            ratio_fig.tight_layout()
            ratio_fig.savefig(ratio_hist_file)
            
            # 4. Save statistics as text file
            stats_file = os.path.join(folder_path, f"{folder_name}_statistics.txt")
            with open(stats_file, 'w') as f:
                f.write(f"File: {self.file_path if hasattr(self, 'file_path') else 'Unknown'}\n")
                f.write(f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                f.write(f"Number of Pulses: {len(self.pulses) if hasattr(self, 'pulses') else 0}\n")
                f.write(f"Number of Periods: {len(self.current_periods)}\n\n")
                
                # Period statistics
                f.write(f"Period Statistics (ms):\n")
                f.write(f"  Mean: {np.mean(durations):.2f}\n")
                f.write(f"  Median: {np.median(durations):.2f}\n")
                f.write(f"  Mode: {mode_value if 'mode_value' in locals() else 0:.2f}\n")
                f.write(f"  Std Dev: {np.std(durations):.2f}\n")
                f.write(f"  Min: {np.min(durations):.2f}\n")
                f.write(f"  Max: {np.max(durations):.2f}\n\n")
                
                # Ratio statistics
                f.write(f"Pulse Ratio Statistics:\n")
                f.write(f"  Mean: {np.mean(ratios):.4f}\n")
                f.write(f"  Median: {np.median(ratios):.4f}\n")
                f.write(f"  Mode: {mode_value if 'mode_value' in locals() else 0:.4f}\n")
                f.write(f"  Std Dev: {np.std(ratios):.4f}\n")
                f.write(f"  Min: {np.min(ratios):.4f}\n")
                f.write(f"  Max: {np.max(ratios):.4f}\n\n")
                
                # Processing information
                threshold_type = "Absolute" if hasattr(self, 'using_absolute_threshold') and self.using_absolute_threshold else "Relative"
                threshold_value = self.abs_threshold if hasattr(self, 'abs_threshold') and hasattr(self, 'using_absolute_threshold') and self.using_absolute_threshold else self.rel_threshold if hasattr(self, 'rel_threshold') else 0
                f.write(f"Processing Information:\n")
                f.write(f"  Inversion Count: {self.inversion_count if hasattr(self, 'inversion_count') else 0}\n")
                f.write(f"  Threshold Type: {threshold_type}\n")
                f.write(f"  Threshold Value: {threshold_value:.3f}\n")
                f.write(f"  Sample Rate: {self.sample_rate} Hz\n")
                f.write(f"  Total Duration: {self.total_frames / self.sample_rate:.2f} seconds\n")
            
            # 5. Save the WAV file
            wav_file = os.path.join(folder_path, f"{folder_name}.wav")
            wavfile.write(wav_file, self.sample_rate, self.wav_data)
            
            QMessageBox.information(self, "Save Successful", 
                                  f"Results saved to folder:\n{folder_path}\n\nFiles created:\n"
                                  f"- {os.path.basename(csv_file)}\n"
                                  f"- {os.path.basename(period_hist_file)}\n"
                                  f"- {os.path.basename(ratio_hist_file)}\n"
                                  f"- {os.path.basename(stats_file)}\n"
                                  f"- {os.path.basename(wav_file)}")
                                  
        except Exception as e:
            QMessageBox.critical(self, "Error Saving Results", 
                               f"Failed to save results: {str(e)}")
            return  # Don't proceed to next file if there was an error
        
        # Check if there are more files in the queue
        if hasattr(self, 'file_queue') and len(self.file_queue) > 1 and self.current_file_index < len(self.file_queue) - 1:
            # Move to the next file
            self.current_file_index += 1
            next_file = self.file_queue[self.current_file_index]
            
            # Add information about next file
            QMessageBox.information(self, "Next File", 
                                  f"Moving to next file ({self.current_file_index + 1}/{len(self.file_queue)}):\n{os.path.basename(next_file)}")
            
            # Load the next file
            self.load_wav_file(next_file)
    
    def reset_application(self):
        """Reset the application to its initial state with the current file."""
        if not hasattr(self, 'original_wav_data') or self.original_wav_data is None:
            QMessageBox.warning(self, "Cannot Reset", "No file has been loaded yet.")
            return
            
        # Confirm with user
        reply = QMessageBox.question(self, 'Reset Application', 
                                   'This will clear all pulses and processing. Continue?',
                                   QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        
        if reply == QMessageBox.Yes:
            # Restore original data
            self.wav_data = self.original_wav_data.copy()
            
            # Reset all processing variables
            self.abs_data = None
            self.smoothed_data = None
            self.pulses = []
            self.skips = []
            
            # Reset view to initial state
            self.view_start = 0
            self.view_range = min(self.sample_rate * 2, self.total_frames)  # View first 2 seconds
            
            # Clear command history
            self.command_history = []
            self.command_index = 0
            
            # Update the plot
            self.update_plot()
            
            QMessageBox.information(self, "Reset Complete", "Application has been reset.")
            
    def show_help(self):
        """Show a help dialog with all available controls."""
        help_text = """
        <html>
        <h2>Katydid Analyzer Controls</h2>
        
        <h3>Navigation Controls:</h3>
        <ul>
            <li><b>A / D:</b> Move left / right</li>
            <li><b>W / S:</b> Zoom in / out</li>
            <li><b>Mouse Drag:</b> Pan the view</li>
            <li><b>Mouse Wheel:</b> Zoom in/out at cursor position</li>
        </ul>
        
        <h3>Signal Processing:</h3>
        <ul>
            <li><b>R:</b> Invert values (flip positive/negative)</li>
            <li><b>G:</b> Apply smoothing</li>
            <li><b>Up/Down Arrows:</b> Adjust threshold</li>
            <li><b>Tab:</b> Toggle between absolute/relative threshold</li>
        </ul>
        
        <h3>Pulse Detection:</h3>
        <ul>
            <li><b>Y:</b> Detect pulses in current view</li>
            <li><b>T:</b> Analyze pulse periods</li>
            <li><b>O:</b> Add manual pulse at selection</li>
            <li><b>P:</b> Delete pulses in selection</li>
        </ul>
        
        <h3>Selection Controls:</h3>
        <ul>
            <li><b>Click and Drag:</b> Select a region</li>
            <li><b>Double Click:</b> Clear selection</li>
        </ul>
        
        <h3>File Operations:</h3>
        <ul>
            <li><b>Analyze Periods:</b> Analyze and optionally save pulse period data</li>
            <li><b>Reset:</b> Reset to original waveform and clear all pulses</li>
        </ul>
        
        </html>
        """
        
        # Create dialog
        help_dialog = QDialog(self)
        help_dialog.setWindowTitle("Katydid Analyzer Help")
        help_dialog.setMinimumSize(500, 600)
        
        # Create layout
        layout = QVBoxLayout()
        
        # Add help text
        help_label = QLabel(help_text)
        help_label.setTextFormat(Qt.RichText)
        help_label.setWordWrap(True)
        help_label.setOpenExternalLinks(True)
        
        # Add to scrollable area
        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(help_label)
        
        # Add to layout
        layout.addWidget(scroll_area)
        
        # Add close button
        close_button = QPushButton("Close")
        close_button.clicked.connect(help_dialog.close)
        layout.addWidget(close_button)
        
        # Set layout and show dialog
        help_dialog.setLayout(layout)
        help_dialog.exec_()
            

    def _detect_skips_in_csv(self):
        # Calculate intervals between pulses
        intervals = []
        for i in range(1, len(self.csv_data)):
            interval = self.csv_data[i]['start'] - self.csv_data[i-1]['start']
            intervals.append(interval)
        
        if len(intervals) < 4:  # Need at least 4 intervals to determine pattern
            QMessageBox.warning(self, "Skip Detection", "Not enough pulses to determine pattern.")
            return
            
        # Find the mode of intervals using histogram
        hist, bin_edges = np.histogram(intervals, bins=50)
        mode_idx = np.argmax(hist)
        mode_interval = (bin_edges[mode_idx] + bin_edges[mode_idx + 1]) / 2
        
        # Check if we have a clear bimodal distribution
        short_durations = [d for d in intervals if d < mode_interval]
        long_durations = [d for d in intervals if d >= mode_interval]
        
        if len(short_durations) > 3 and len(long_durations) > 3:
            # Likely double pulse pattern
            self.pulse_type = "double"
            short_mean = np.mean(short_durations)
            long_mean = np.mean(long_durations)
            
            # Detect skips as abnormal long intervals
            self.skips = []
            for i, item in enumerate(self.csv_data):
                if item['interval'] > long_mean * 1.3 and item['interval'] < long_mean * 2.5:
                    self.skips.append(item)
        else:  # Single pulse
            # Check each interval against the mode
            for i, interval in enumerate(intervals):
                if not (min(mode_interval * 0.8, mode_interval * 1.2) <= interval <= max(mode_interval * 0.8, mode_interval * 1.2)):
                    # This is a potential skip
                    self.skips.append({
                        'position': self.csv_data[i]['start'],
                        'type': 'irregular_interval',
                        'interval': interval,
                        'notes': f"Irregular interval at {self.csv_data[i]['start']:.2f}ms"
                    })
    
    def populate_results_table(self):
        if not self.skips:
            QMessageBox.information(self, "Skip Detection", "No skips detected in the recording.")
            return

        if not hasattr(self, 'results_window') or self.results_window is None:
            self.results_window = QDialog(self)
            self.results_window.setWindowTitle("Skip Detection Results")
            self.results_window.setMinimumSize(800, 600)  # Increased size
            
            # Create layout
            layout = QVBoxLayout(self.results_window)
            
            # Add explanation label
            info_label = QLabel("The table below shows detected skips and their characteristics:")
            info_label.setWordWrap(True)
            layout.addWidget(info_label)
            
            if self.skips:
                skip_label = QLabel(f"Found {len(self.skips)} potential skips:")
            else:
                skip_label = QLabel("No skips detected in the data.")
            layout.addWidget(skip_label)
            
            # Create table for results
            if self.skips:
                results_table = QTableWidget()
                results_table.setColumnCount(4)
                
                # Set column headers
                results_table.setHorizontalHeaderLabels(["Position (ms)", "Interval (ms)", "Type", ""])
                
                results_table.setStyleSheet("""
                    QTableWidget {
                        background-color: white;
                        border: 1px solid #ddd;
                    }
                    QHeaderView::section {
                        background-color: #f0f9fa;
                        padding: 4px;
                        border: 1px solid #ddd;
                        font-weight: bold;
                    }
                """)
                
                # Add click handler to focus on graph
                results_table.itemClicked.connect(self.on_skip_selected)
                
                layout.addWidget(results_table)
                
                # Populate table with skip data
                for row, skip in enumerate(self.skips):
                    results_table.insertRow(row)
                    
                    # Position
                    if self.pulse_type == "csv":
                        position = skip['position']
                    else:
                        position = skip['position'] / self.sample_rate * 1000
                    results_table.setItem(row, 0, QTableWidgetItem(f"{position:.2f}"))
                    
                    # Interval
                    results_table.setItem(row, 1, QTableWidgetItem(f"{skip['interval']:.2f}"))
                    
                    # Type
                    results_table.setItem(row, 2, QTableWidgetItem(skip['type']))
                    
                    # Add green button
                    button = QPushButton("Go To")
                    button.setStyleSheet("""
                        QPushButton {
                            background-color: #00ff00;
                            color: black;
                            border: none;
                            padding: 4px;
                        }
                        QPushButton:hover {
                            background-color: #00cc00;
                        }
                    """)
                    button.clicked.connect(lambda _, r=row: self.go_to_skip(r))
                    results_table.setCellWidget(row, 3, button)
                
                results_table.resizeColumnsToContents()
            
            # Add save button
            save_button = QPushButton("Save Results")
            save_button.clicked.connect(self.save_results)
            layout.addWidget(save_button)
            
            # Show the window
            self.results_window.setLayout(layout)
            self.results_window.show()
        else:
            self.results_window.show()

    def go_to_skip(self, row):
        """Center the graph view on the selected skip."""
        if row < len(self.skips):
            skip = self.skips[row]
            
            # Center the view around the skip position
            if self.pulse_type == "csv":
                position = skip['position']
                self.view_start = max(0, int(position - self.view_range / 2))
            else:
                position = skip['position']  # Position is already in samples
                self.view_start = max(0, int(position - self.view_range / 2))
            
            # Load the data chunk at this position to ensure the graph displays properly
            self.load_chunk(int(self.view_start))
            self.update_plot()
            
            # Highlight the selected skip in the table
            self.results_table.selectRow(row)
    
    def on_skip_selected(self, item):
        """Handle skip selection in the results table."""
        if item.column() == 0:  # Only respond to row number clicks
            row = item.row()
            if row < len(self.skips):
                skip = self.skips[row]
                
                # Center the view around the skip position
                if self.pulse_type == "csv":
                    position = skip['position']
                    self.view_start = max(0, position - self.view_range / 2)
                else:
                    position = skip['position']  # Position is already in samples
                    self.view_start = max(0, position - self.view_range / 2)
                
                # Load the data chunk at this position to ensure the graph displays properly
                self.load_chunk(int(self.view_start))
                self.update_plot()
    
    def save_results(self):
        """Save analysis results to a folder with user-specified name."""
        # Check if we have periods to save
        if not hasattr(self, 'current_periods') or not self.current_periods:
            QMessageBox.warning(self, "No Data to Save", "Please analyze pulse periods first (press T) before saving.")
            return
            
        # Ask user for a folder name
        folder_name, ok = QInputDialog.getText(self, "Save Results", "What do you want your folder name to be?", 
                                          QLineEdit.Normal, "")
        if not ok or not folder_name:
            return
            
        # Ask user for a directory to save the folder in
        save_dir = QFileDialog.getExistingDirectory(self, "Select Directory to Save Folder")
        if not save_dir:  # User canceled
            return
            
        # Create the folder
        folder_path = os.path.join(save_dir, folder_name)
        try:
            os.makedirs(folder_path, exist_ok=True)
        except Exception as e:
            QMessageBox.critical(self, "Error Creating Folder", f"Failed to create folder: {str(e)}")
            return
            
        try:
            # Create timestamp for unique filenames if needed
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            
            # Sort pulses by position to ensure proper ordering
            sorted_pulses = sorted(self.pulses, key=lambda p: p['position'])
            
            # 1. Save the table as CSV with amplitude information
            csv_file = os.path.join(folder_path, f"{folder_name}_table.csv")
            with open(csv_file, 'w', newline='') as f:
                writer = csv.writer(f)
                writer.writerow(['Period', 'Duration (ms)', 'Pulse Ratio', 'Pulse1 Time (ms)', 'Pulse1 Amplitude', 'Pulse2 Time (ms)', 'Pulse2 Amplitude'])
                
                for period in self.current_periods:
                    # Get the pulses for this period
                    period_idx = period['index'] - 1  # Convert to 0-based index
                    
                    # Get pulse times and amplitudes from sorted pulses
                    if period_idx + 2 < len(sorted_pulses):
                        pulse1 = sorted_pulses[period_idx]
                        pulse2 = sorted_pulses[period_idx + 1]
                        pulse3 = sorted_pulses[period_idx + 2]
                    else:
                        pulse1 = pulse2 = pulse3 = {'time': 0, 'amplitude': 0}
                    
                    writer.writerow([period['index'], 
                                    period['duration'] if 'duration' in period else 0,
                                    period['ratio'] if 'ratio' in period else 0,
                                    pulse1['time'] if 'time' in pulse1 else 0,
                                    pulse1['amplitude'] if 'amplitude' in pulse1 else 0,
                                    pulse2['time'] if 'time' in pulse2 else 0,
                                    pulse2['amplitude'] if 'amplitude' in pulse2 else 0])
            
            # 2. Save the period duration histogram
            period_hist_file = os.path.join(folder_path, f"{folder_name}_period_histogram.png")
            
            # Create figure for period histogram
            period_fig = Figure(figsize=(8, 6))
            period_canvas = FigureCanvas(period_fig)
            period_ax = period_fig.add_subplot(111)
            
            # Extract duration data
            durations = [p['duration'] for p in self.current_periods]
            
            # Calculate mode of durations
            if durations:
                # Create bins for histogram
                hist, bin_edges = np.histogram(durations, bins=30)
                # Find the bin with the highest count
                mode_bin_index = np.argmax(hist)
                # Get the mode range
                mode_range = (bin_edges[mode_bin_index], bin_edges[mode_bin_index + 1])
                mode_value = (mode_range[0] + mode_range[1]) / 2
                
                # Plot histogram of durations
                n, bins, patches = period_ax.hist(durations, bins=30, alpha=0.7, color='green')
                
                # Highlight the mode bin
                for i, patch in enumerate(patches):
                    if i == mode_bin_index:
                        patch.set_facecolor('red')  # Highlight the mode bin
            
                # Add a vertical line at the mode
                period_ax.axvline(x=mode_value, color='red', linestyle='--', linewidth=2)
                period_ax.text(mode_value, max(n)*0.9, f'Mode: {mode_value:.2f} ms', 
                            color='red', fontweight='bold', ha='right')
            else:
                period_ax.text(0.5, 0.5, 'No data available', ha='center', va='center', transform=period_ax.transAxes)
            
            period_ax.set_xlabel('Period Duration (ms)')
            period_ax.set_ylabel('Frequency')
            period_ax.set_title('Distribution of Period Durations (Mode Highlighted)')
            period_ax.grid(True)
            period_ax.legend()
            period_ax.legend(loc='upper right')  # Specify a fixed location for the legend
            period_fig.tight_layout()
            period_fig.savefig(period_hist_file)
            
            # 3. Save the ratio histogram
            ratio_hist_file = os.path.join(folder_path, f"{folder_name}_ratio_histogram.png")
            ratio_fig = Figure(figsize=(8, 6))
            ratio_canvas = FigureCanvas(ratio_fig)
            ratio_ax = ratio_fig.add_subplot(111)
            
            # Extract ratio data
            ratios = [p['ratio'] for p in self.current_periods]
            
            # Calculate mode of ratios
            if ratios:
                # Create bins for histogram
                hist, bin_edges = np.histogram(ratios, bins=30)
                # Find the bin with the highest count
                mode_bin_index = np.argmax(hist)
                # Get the mode range
                mode_range = (bin_edges[mode_bin_index], bin_edges[mode_bin_index + 1])
                mode_value = (mode_range[0] + mode_range[1]) / 2
                
                # Plot histogram of ratios
                n, bins, patches = ratio_ax.hist(ratios, bins=30, alpha=0.7, color='blue')
                
                # Highlight the mode bin
                for i, patch in enumerate(patches):
                    if i == mode_bin_index:
                        patch.set_facecolor('red')  # Highlight the mode bin
            
                # Add a vertical line at the mode
                ratio_ax.axvline(x=mode_value, color='red', linestyle='--', linewidth=2)
                ratio_ax.text(mode_value, max(n)*0.9, f'Mode: {mode_value:.4f}', 
                           color='red', fontweight='bold', ha='right')
            else:
                ratio_ax.text(0.5, 0.5, 'No data available', ha='center', va='center', transform=ratio_ax.transAxes)
            
            ratio_ax.set_xlabel('Pulse Ratio (time between pulses 1-2 / period duration)')
            ratio_ax.set_ylabel('Frequency')
            ratio_ax.set_title('Distribution of Pulse Ratios (Mode Highlighted)')
            ratio_ax.grid(True)
            ratio_ax.legend()
            ratio_ax.legend(loc='upper right')  # Specify a fixed location for the legend
            ratio_fig.tight_layout()
            ratio_fig.savefig(ratio_hist_file)
            
            # 4. Save statistics as text file
            stats_file = os.path.join(folder_path, f"{folder_name}_statistics.txt")
            with open(stats_file, 'w') as f:
                f.write(f"File: {self.file_path if hasattr(self, 'file_path') else 'Unknown'}\n")
                f.write(f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
                f.write(f"Number of Pulses: {len(self.pulses) if hasattr(self, 'pulses') else 0}\n")
                f.write(f"Number of Periods: {len(self.current_periods)}\n\n")
                
                # Period statistics
                f.write(f"Period Statistics (ms):\n")
                f.write(f"  Mean: {np.mean(durations):.2f}\n")
                f.write(f"  Median: {np.median(durations):.2f}\n")
                f.write(f"  Mode: {mode_value if 'mode_value' in locals() else 0:.2f}\n")
                f.write(f"  Std Dev: {np.std(durations):.2f}\n")
                f.write(f"  Min: {np.min(durations):.2f}\n")
                f.write(f"  Max: {np.max(durations):.2f}\n\n")
                
                # Ratio statistics
                f.write(f"Pulse Ratio Statistics:\n")
                f.write(f"  Mean: {np.mean(ratios):.4f}\n")
                f.write(f"  Median: {np.median(ratios):.4f}\n")
                f.write(f"  Mode: {mode_value if 'mode_value' in locals() else 0:.4f}\n")
                f.write(f"  Std Dev: {np.std(ratios):.4f}\n")
                f.write(f"  Min: {np.min(ratios):.4f}\n")
                f.write(f"  Max: {np.max(ratios):.4f}\n\n")
                
                # Processing information
                threshold_type = "Absolute" if hasattr(self, 'using_absolute_threshold') and self.using_absolute_threshold else "Relative"
                threshold_value = self.abs_threshold if hasattr(self, 'abs_threshold') and hasattr(self, 'using_absolute_threshold') and self.using_absolute_threshold else self.rel_threshold if hasattr(self, 'rel_threshold') else 0
                f.write(f"Processing Information:\n")
                f.write(f"  Inversion Count: {self.inversion_count if hasattr(self, 'inversion_count') else 0}\n")
                f.write(f"  Threshold Type: {threshold_type}\n")
                f.write(f"  Threshold Value: {threshold_value:.3f}\n")
            
            QMessageBox.information(self, "Save Successful", 
                                  f"Results saved to folder:\n{folder_path}\n\nFiles created:\n"
                                  f"- {os.path.basename(csv_file)}\n"
                                  f"- {os.path.basename(period_hist_file)}\n"
                                  f"- {os.path.basename(ratio_hist_file)}\n"
                                  f"- {os.path.basename(stats_file)}")
                                  
        except Exception as e:
            QMessageBox.critical(self, "Error Saving Results", 
                               f"Failed to save results: {str(e)}")
            return  # Don't proceed to next file if there was an error
        
        # Check if there are more files in the queue
        if hasattr(self, 'file_queue') and len(self.file_queue) > 1 and self.current_file_index < len(self.file_queue) - 1:
            # Move to the next file
            self.current_file_index += 1
            next_file = self.file_queue[self.current_file_index]
            
            # Add information about next file
            QMessageBox.information(self, "Next File", 
                                  f"Moving to next file ({self.current_file_index + 1}/{len(self.file_queue)}):\n{os.path.basename(next_file)}")
            
            # Load the next file
            self.load_wav_file(next_file)

        # 2. Save the period duration histogram
        period_hist_file = os.path.join(folder_path, f"{folder_name}_period_histogram.png")
        
        # Create figure for period histogram
        period_fig = Figure(figsize=(8, 6))
        period_canvas = FigureCanvas(period_fig)
        period_ax = period_fig.add_subplot(111)
        
        # Extract duration data
        durations = [p['duration'] for p in self.current_periods]
        
        # Calculate mode of durations
        if durations:
            # Create bins for histogram
            hist, bin_edges = np.histogram(durations, bins=30)
            # Find the bin with the highest count
            mode_bin_index = np.argmax(hist)
            # Get the mode range
            mode_range = (bin_edges[mode_bin_index], bin_edges[mode_bin_index + 1])
            mode_value = (mode_range[0] + mode_range[1]) / 2
            
            # Plot histogram of durations
            n, bins, patches = period_ax.hist(durations, bins=30, alpha=0.7, color='green')
            
            # Highlight the mode bin
            for i, patch in enumerate(patches):
                if i == mode_bin_index:
                    patch.set_facecolor('red')  # Highlight the mode bin
            
            # Add a vertical line at the mode
            period_ax.axvline(x=mode_value, color='red', linestyle='--', linewidth=2)
            period_ax.text(mode_value, max(n)*0.9, f'Mode: {mode_value:.2f} ms', 
                        color='red', fontweight='bold', ha='right')
        else:
            period_ax.text(0.5, 0.5, 'No data available', ha='center', va='center', transform=period_ax.transAxes)
        
        period_ax.set_xlabel('Period Duration (ms)')
        period_ax.set_ylabel('Frequency')
        period_ax.set_title('Distribution of Period Durations (Mode Highlighted)')
        period_ax.grid(True)
        period_ax.legend()
        period_ax.legend(loc='upper right')  # Specify a fixed location for the legend
        period_fig.tight_layout()
        period_fig.savefig(period_hist_file)
        
        # 3. Save the ratio histogram
        ratio_hist_file = os.path.join(folder_path, f"{folder_name}_ratio_histogram.png")
        ratio_fig = Figure(figsize=(8, 6))
        ratio_canvas = FigureCanvas(ratio_fig)
        ratio_ax = ratio_fig.add_subplot(111)
        
        # Extract ratio data
        ratios = [p['ratio'] for p in self.current_periods]
        
        # Calculate mode of ratios
        if ratios:
            # Create bins for histogram
            hist, bin_edges = np.histogram(ratios, bins=30)
            # Find the bin with the highest count
            mode_bin_index = np.argmax(hist)
            # Get the mode range
            mode_range = (bin_edges[mode_bin_index], bin_edges[mode_bin_index + 1])
            mode_value = (mode_range[0] + mode_range[1]) / 2
            
            # Plot histogram of ratios
            n, bins, patches = ratio_ax.hist(ratios, bins=30, alpha=0.7, color='blue')
            
            # Highlight the mode bin
            for i, patch in enumerate(patches):
                if i == mode_bin_index:
                    patch.set_facecolor('red')  # Highlight the mode bin
            
            # Add a vertical line at the mode
            ratio_ax.axvline(x=mode_value, color='red', linestyle='--', linewidth=2)
            ratio_ax.text(mode_value, max(n)*0.9, f'Mode: {mode_value:.4f}', 
                       color='red', fontweight='bold', ha='right')
        else:
            ratio_ax.text(0.5, 0.5, 'No data available', ha='center', va='center', transform=ratio_ax.transAxes)
        
        ratio_ax.set_xlabel('Pulse Ratio (time between pulses 1-2 / period duration)')
        ratio_ax.set_ylabel('Frequency')
        ratio_ax.set_title('Distribution of Pulse Ratios (Mode Highlighted)')
        ratio_ax.grid(True)
        ratio_ax.legend()
        ratio_ax.legend(loc='upper right')  # Specify a fixed location for the legend
        ratio_fig.tight_layout()
        ratio_fig.savefig(ratio_hist_file)
        
        # 4. Save statistics as text file
        stats_file = os.path.join(folder_path, f"{folder_name}_statistics.txt")
        with open(stats_file, 'w') as f:
            f.write(f"File: {self.file_path if hasattr(self, 'file_path') else 'Unknown'}\n")
            f.write(f"Analysis Date: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            f.write(f"Number of Pulses: {len(self.pulses) if hasattr(self, 'pulses') else 0}\n")
            f.write(f"Number of Periods: {len(self.current_periods)}\n\n")
            
            # Period statistics
            f.write(f"Period Statistics (ms):\n")
            f.write(f"  Mean: {np.mean(durations):.2f}\n")
            f.write(f"  Median: {np.median(durations):.2f}\n")
            f.write(f"  Mode: {mode_value if 'mode_value' in locals() else 0:.2f}\n")
            f.write(f"  Std Dev: {np.std(durations):.2f}\n")
            f.write(f"  Min: {np.min(durations):.2f}\n")
            f.write(f"  Max: {np.max(durations):.2f}\n\n")
            
            # Ratio statistics
            f.write(f"Pulse Ratio Statistics:\n")
            f.write(f"  Mean: {np.mean(ratios):.4f}\n")
            f.write(f"  Median: {np.median(ratios):.4f}\n")
            f.write(f"  Mode: {mode_value if 'mode_value' in locals() else 0:.4f}\n")
            f.write(f"  Std Dev: {np.std(ratios):.4f}\n")
            f.write(f"  Min: {np.min(ratios):.4f}\n")
            f.write(f"  Max: {np.max(ratios):.4f}\n\n")
            
            # Processing information
            threshold_type = "Absolute" if hasattr(self, 'using_absolute_threshold') and self.using_absolute_threshold else "Relative"
            threshold_value = self.abs_threshold if hasattr(self, 'abs_threshold') and hasattr(self, 'using_absolute_threshold') and self.using_absolute_threshold else self.rel_threshold if hasattr(self, 'rel_threshold') else 0
            f.write(f"Processing Information:\n")
            f.write(f"  Inversion Count: {self.inversion_count if hasattr(self, 'inversion_count') else 0}\n")
            f.write(f"  Threshold Type: {threshold_type}\n")
            f.write(f"  Threshold Value: {threshold_value:.3f}\n")

        # Show success message and handle next file
        try:
            QMessageBox.information(self, "Save Successful", 
                f"Results saved to folder:\n{folder_path}\n\nFiles created:\n"
                f"- {os.path.basename(csv_file)}\n"
                f"- {os.path.basename(period_hist_file)}\n"
                f"- {os.path.basename(ratio_hist_file)}\n"
                f"- {os.path.basename(stats_file)}")
            
            # Check if there are more files in the queue
            if hasattr(self, 'file_queue') and len(self.file_queue) > 1 and self.current_file_index < len(self.file_queue) - 1:
                # Move to the next file
                self.current_file_index += 1
                next_file = self.file_queue[self.current_file_index]
                
                # Add information about next file
                QMessageBox.information(self, "Next File", 
                    f"Moving to next file ({self.current_file_index + 1}/{len(self.file_queue)}):\n{os.path.basename(next_file)}")
                
                # Load the next file
                self.load_wav_file(next_file)
        except Exception as e:
            QMessageBox.critical(self, "Error Saving Results", 
                f"Failed to save results: {str(e)}")

    
    def show_status_message(self, message, duration=1500):
        """Show a status message overlay on the plot"""
        # Remove any existing status message
        if self.status_message:
            try:
                self.status_message.remove()
            except:
                pass
        
        # Create new status message
        if hasattr(self, 'ax') and self.ax:
            self.status_message = self.ax.text(
                0.5, 0.95, 
                message, 
                transform=self.ax.transAxes,
                ha='center', va='top',
                bbox=dict(
                    facecolor='black', 
                    alpha=0.7,
                    edgecolor='white',
                    boxstyle='round,pad=0.5'
                ),
                color='#00ff00',
                fontsize=10,
                weight='bold'
            )
            self.canvas.draw_idle()
            
            # Set timer to remove the message
            if self.status_timer:
                self.status_timer.stop()
            
            # Use QTimer to clear the message after duration
            from PyQt5.QtCore import QTimer
            self.status_timer = QTimer()
            self.status_timer.setSingleShot(True)
            self.status_timer.timeout.connect(self._clear_status_message)
            self.status_timer.start(duration)
    
    def _clear_status_message(self):
        """Clear the status message"""
        if self.status_message:
            try:
                self.status_message.remove()
                self.status_message = None
                self.canvas.draw_idle()
            except:
                pass

from scipy.io import wavfile  # Import here to avoid potential circular imports

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = KatydidAnalysisApp()
    window.show()
    sys.exit(app.exec_())   
