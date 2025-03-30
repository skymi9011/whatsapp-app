import sys
import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import quote
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                            QLabel, QPushButton, QFileDialog, QProgressBar, QTextEdit,
                            QTableWidget, QTableWidgetItem, QHeaderView, QMessageBox, QLineEdit,
                            QFrame, QScrollArea , QTabWidget, QSizePolicy, QGroupBox, QSpacerItem)
from PyQt5.QtGui import QPixmap, QIcon, QFont, QColor, QPalette, QFontDatabase, QDrag, QIntValidator
from PyQt5.QtCore import Qt, QThread, pyqtSignal, QSize, QUrl, QMimeData
import requests

# Define brand colors with adjusted shades
PRIMARY_COLOR = "#1976D2"  # Deeper blue
SECONDARY_COLOR = "#00BFA5"  # Teal
BG_COLOR = "#F8F9FA"  # Lighter background
TEXT_COLOR = "#2D3748"  # Darker text for better contrast
SUCCESS_COLOR = "#4CAF50"  # Success green
WARNING_COLOR = "#FF9800"  # Warning orange
ERROR_COLOR = "#F44336"  # Error red
CARD_SHADOW = "rgba(0, 0, 0, 0.1)"  # Shadow color for cards

class WhatsAppSender(QThread):
    update_progress = pyqtSignal(int, str, str)  # Added status type parameter
    finished = pyqtSignal(bool, str)

    def __init__(self, excel_file, delay, parent=None):
        super().__init__(parent)
        self.excel_file = excel_file
        self.delay = delay
        # Default column names
        self.message_column = "Message"
        self.name_column = "Customers Name"
        self.number_column = "Whatsapp Number"
        self.is_running = True

    def stop(self):
        self.is_running = False

    def run(self):
        try:
            # Read Excel file
            df = pd.read_excel(self.excel_file)
            if df[self.number_column].dtype != 'object':
                df[self.number_column] = df[self.number_column].astype(str)
            df[self.number_column] = df[self.number_column].str.replace(r'\D+', '', regex=True)  # Clean numbers
            
            # Step 1: Open browser in visible mode for QR code scanning
            chrome_profile_path = os.path.join(os.environ['USERPROFILE'], 'AppData', 'Local', 'Google', 'Chrome', 'User Data', 'Default')
            options = webdriver.ChromeOptions()
            options.add_argument(f'user-data-dir={chrome_profile_path}')
            driver = webdriver.Chrome(options=options)
            driver.get('https://web.whatsapp.com/')
            
            try:
                # Check if the user is logged in
                WebDriverWait(driver, 10).until(
                    EC.presence_of_element_located((By.XPATH, '//div[@role="textbox"][@contenteditable="true"]'))
                )
                self.update_progress.emit(0, "Logged into WhatsApp Web successfully.", "success")
            except Exception:
                # If not logged in, wait for QR code scan
                self.update_progress.emit(0, "Please scan the QR code to log in.", "info")
                try:
                    WebDriverWait(driver, 60).until(
                        EC.presence_of_element_located((By.XPATH, '//div[@role="textbox"][@contenteditable="true"]'))
                    )
                    self.update_progress.emit(0, "Logged into WhatsApp Web successfully.", "success")
                except Exception as e:
                    driver.quit()
                    self.finished.emit(False, f"Failed to log in: {str(e)}")
                    return
            
            # Step 2: Close the visible browser and reopen in headless mode
            driver.quit()
            options.add_argument('--headless')  # Enable headless mode
            options.add_argument('--disable-gpu')
            options.add_argument('--no-sandbox')
            options.add_argument('--disable-dev-shm-usage')
            driver = webdriver.Chrome(options=options)
            
            total_count = len(df)
            success_count = 0
            
            for index, row in df.iterrows():
                if not self.is_running:
                    driver.quit()
                    self.finished.emit(False, "Process was stopped by user")
                    return
                
                try:
                    # Define name and number before using them
                    name = row[self.name_column]
                    number = row[self.number_column]
                    message = str(row[self.message_column]).replace('{name}', name)
                    
                    # Generate WhatsApp URL
                    url = f'https://web.whatsapp.com/send?phone={number}&text={quote(message)}'
                    driver.get(url)
                    
                    # Wait for the message input box to load
                    try:
                        WebDriverWait(driver, 15).until(
                            EC.presence_of_element_located((By.XPATH, '//div[@role="textbox"][@contenteditable="true"]'))
                        )
                        
                        # Click send button
                        send_button = WebDriverWait(driver, 15).until(
                            EC.element_to_be_clickable((By.XPATH, '//span[@data-icon="send"]'))
                        )
                        send_button.click()
                        
                        success_count += 1
                        status_msg = f"Sent to {name} ({number})"
                        status_type = "success"
                    except Exception as e:
                        status_msg = f"Failed to send to {name} ({number}): {str(e)}"
                        status_type = "error"
                    
                    progress_percent = int((index + 1) / total_count * 100)
                    self.update_progress.emit(progress_percent, status_msg, status_type)
                    time.sleep(self.delay)  # Wait between messages
                    
                except Exception as e:
                    self.update_progress.emit(
                        int((index + 1) / total_count * 100), 
                        f"Error processing {name} ({number}): {str(e)}",
                        "error"
                    )
            
            driver.quit()
            self.finished.emit(True, f"Completed sending messages. Success: {success_count}/{total_count}")
            
        except Exception as e:
            self.finished.emit(False, f"Process failed: {str(e)}")
class StyleSheet:
    @staticmethod
    def get_stylesheet():
        return f"""
        /* Base Styling */
        

        QMainWindow, QWidget {{
            background-color: #F8FAFC;
            color: #1E293B;
            font-family: 'Segoe UI', -apple-system, system-ui;
            font-size: 13px;
        }}

        /* Card Containers */
        QFrame#CardFrame {{
            background: white;
            border-radius: 8px;
            border: 1px solid #E2E8F0;
            padding: 20px;
            margin: 4px 0;
        }}

        /* Typography Hierarchy */
        QLabel#HeaderTitle {{
            font-size: 20px;
            font-weight: 600;
            color: #0F172A;
            padding: 8px 0;
        }}

        QLabel#SectionTitle {{
            font-size: 14px;
            font-weight: 500;
            color: #475569;
            letter-spacing: 0.5px;
            text-transform: uppercase;
            margin-bottom: 12px;
        }}

        /* Buttons - Modern Minimal */
        QPushButton {{
            border: 1px solid #CBD5E1;
            border-radius: 6px;
            padding: 8px 16px;
            background: white;
            color: #334155;
            font-weight: 500;
            min-width: 100px;
        }}

        QPushButton:hover {{
            background: #F1F5F9;
            border-color: #94A3B8;
        }}

        QPushButton:pressed {{
            background: #E2E8F0;
        }}

        QPushButton#PrimaryButton {{
            background: #2563EB;
            color: white;
            border-color: #2563EB;
        }}

        QPushButton#PrimaryButton:hover {{
            background: #1D4ED8;
        }}

        QPushButton#DangerButton {{
            background: #DC2626;
            color: white;
            border-color: #DC2626;
        }}

        /* Input Fields */
        QLineEdit, QTextEdit {{
            border: 1px solid #CBD5E1;
            border-radius: 6px;
            padding: 8px 12px;
            background: white;
            selection-background-color: #BFDBFE;
            font-size: 13px;
        }}

        QLineEdit:focus, QTextEdit:focus {{
            border: 1px solid #2563EB;
            outline: none;
        }}

        /* Progress Bar - Subtle */
        QProgressBar {{
            border: 1px solid #E2E8F0;
            border-radius: 4px;
            height: 18px;
            text-align: center;
        }}

        QProgressBar::chunk {{
            background: #2563EB;
            border-radius: 3px;
        }}

        /* Table - Clean Data Grid */
        QTableWidget {{
            border: none;
            background: white;
            alternate-background-color: #F8FAFC;
        }}

        QHeaderView::section {{
            background: white;
            border: none;
            border-bottom: 2px solid #E2E8F0;
            padding: 8px;
            font-weight: 500;
        }}

        QTableWidget::item {{
            padding: 8px;
            border-bottom: 1px solid #F1F5F9;
        }}
        QLabel#DropLabel {{
            background: transparent;  /* إزالة الخلفية */
            font-size: 16px;  /* حجم الخط */
            font-weight: 500;  /* وزن الخط */
            color: #475569;  /* لون النص */
            text-align: center;  /* محاذاة النص في المنتصف */
        }}
        QLabel#IconLabel {{
            background: transparent;  /* إزالة الخلفية */
            font-size: 48px;  /* حجم الخط للأيقونة */
            color: #475569;  /* لون الأيقونة */
            text-align: center;  /* محاذاة الأيقونة في المنتصف */
        }}
        /* Drop Zone - Modern */
        QFrame#DropArea {{
            background: transparent;  /* إزالة الخلفية */
            border: 2px dashed #2563EB;  /* إطار متقطع باللون الأزرق */
            border-radius: 12px;  /* زوايا مستديرة */
            padding: 24px;  /* مسافة داخلية */
            color: #475569;  /* لون النص */
            font-size: 16px;  /* حجم الخط */
            text-align: center;  /* محاذاة ا
        }}

        QFrame#DropAreaActive {{
    border-color: #1D4ED8;  /* لون الإطار عند السحب */
    background: rgba(29, 78, 216, 0.1); 
        }}

        /* Tabs - Minimal Indicator */
        QTabWidget::pane {{
            border: none;
        }}

        QTabBar::tab {{
            background: transparent;
            padding: 8px 16px;
            border-bottom: 2px solid transparent;
            color: #64748B;
        }}

        QTabBar::tab:selected {{
            color: #2563EB;
            border-bottom: 2px solid #2563EB;
        }}

        /* Status Colors */
        QLabel#StatusSuccess {{ color: #16A34A; }}
        QLabel#StatusWarning {{ color: #F59E0B; }}
        QLabel#StatusError {{ color: #DC2626; }}
        QLabel#StatusInfo {{ color: #2563EB; }}
        """

class LogWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QVBoxLayout(self)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        layout.addWidget(self.log_text)
        
        button_layout = QHBoxLayout()
        button_layout.addStretch()
        
        self.clear_button = QPushButton("Clear Log")
        self.clear_button.setFixedWidth(120)
        self.clear_button.clicked.connect(self.clear_log)
        button_layout.addWidget(self.clear_button)
        
        layout.addLayout(button_layout)
    
    def add_log_entry(self, message, status_type="info"):
        color_map = {
            "success": SUCCESS_COLOR,
            "warning": WARNING_COLOR,
            "error": ERROR_COLOR,
            "info": PRIMARY_COLOR
        }
        
        color = color_map.get(status_type, PRIMARY_COLOR)
        timestamp = time.strftime("%H:%M:%S")
        
        self.log_text.append(f'<span style="color: gray;">[{timestamp}]</span> <span style="color: {color};">{message}</span>')
        self.log_text.ensureCursorVisible()
    
    def clear_log(self):
        self.log_text.clear()


class StatusBar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        layout = QHBoxLayout(self)
        layout.setContentsMargins(20, 20, 20, 20)  # Use equal margins on all sides
        layout.setSpacing(15)
        
        # Add a frame with rounded corners for status bar
        status_frame = QFrame()
        status_frame.setObjectName("CardFrame")
        status_frame_layout = QHBoxLayout(status_frame)
        status_frame_layout.setContentsMargins(10, 5, 10, 5)
        
        self.status_icon = QLabel()
        self.status_icon.setFixedSize(16, 16)
        status_frame_layout.addWidget(self.status_icon)
        
        self.status_label = QLabel("Ready")
        status_frame_layout.addWidget(self.status_label)
        
        status_frame_layout.addStretch()
        
        self.progress_bar = QProgressBar()
        self.progress_bar.setFixedWidth(200)
        self.progress_bar.setValue(0)
        status_frame_layout.addWidget(self.progress_bar)
        
        layout.addWidget(status_frame)
    
    def update_status(self, message, progress=None, status_type="info"):
        icon_map = {
            "success": "✓",
            "warning": "⚠",
            "error": "✗",
            "info": "ℹ"
        }
        
        class_map = {
            "success": "StatusSuccess",
            "warning": "StatusWarning",
            "error": "StatusError",
            "info": "StatusInfo"
        }
        
        self.status_icon.setText(icon_map.get(status_type, "ℹ"))
        self.status_label.setText(message)
        
        # Remove all status classes
        for class_name in class_map.values():
            self.status_label.setObjectName("")
        
        # Set new class
        self.status_label.setObjectName(class_map.get(status_type, "StatusInfo"))
        
        if progress is not None:
            self.progress_bar.setValue(progress)


class FileDropArea(QFrame):
    fileDropped = pyqtSignal(str)
    
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setObjectName("DropArea")
        self.setAcceptDrops(True)
        self.setMinimumHeight(200)  # زيادة ارتفاع المربع
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)  # محاذاة المحتويات في المنتصف
        layout.setContentsMargins(20, 20, 20, 20)  # مسافات داخلية متساوية
        layout.setSpacing(20)  # زيادة المسافة بين العناصر
        
        # أيقونة
        self.icon_label = QLabel()
        self.icon_label.setText("📁")  # استخدام رمز أو صورة
        self.icon_label.setObjectName("IconLabel")
        self.icon_label.setFont(QFont("Arial", 48))  # تكبير حجم الأيقونة
        self.icon_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.icon_label)
        
        # نص
        self.drop_label = QLabel("Drag & Drop Excel File Here or Click to Browse")
        self.drop_label.setObjectName("DropLabel")
        self.drop_label.setFont(QFont("Arial", 16))  # تكبير حجم النص
        self.drop_label.setAlignment(Qt.AlignCenter)
        layout.addWidget(self.drop_label)
        
        # زر التصفح
        self.browse_btn = QPushButton("Browse...")
        self.browse_btn.setFixedWidth(150)
        self.browse_btn.setFixedHeight(40)  # زيادة ارتفاع الزر
        self.browse_btn.clicked.connect(self.browse_file)
        layout.addWidget(self.browse_btn, 0, Qt.AlignCenter)
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setObjectName("DropAreaActive")
            self.style().polish(self) 
    
    def dragLeaveEvent(self, event):
        self.setObjectName("DropArea")
        self.style().polish(self)  # Refresh the style
    
    def dropEvent(self, event):
        urls = event.mimeData().urls()
        if urls and urls[0].toLocalFile():
            file_path = urls[0].toLocalFile()
            if file_path.endswith(('.xlsx', '.xls')):  # Check for valid Excel file extensions
                self.fileDropped.emit(file_path)  # Emit the file path
            else:
                QMessageBox.warning(self, "Invalid File", "Please drop a valid Excel file.")
        self.setObjectName("DropArea")
        self.style().polish(self)  # Refresh the style
    
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select Excel File", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.fileDropped.emit(file_path)


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.sender_thread = None
        self.initUI()
        
    def initUI(self):
        self.setStyleSheet(StyleSheet.get_stylesheet())
        self.setWindowTitle("DeerMedia WhatsApp Messenger")
        self.setMinimumSize(1000, 800)
        self.setStyleSheet(StyleSheet.get_stylesheet())
        
        # Main widget with vertical layout
        main_widget = QWidget()
        main_layout = QVBoxLayout(main_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)
        
        # حقوق الملكية
        copyright_label = QLabel("جميع الحقوق محفوظة لمؤسسة الغزال للإنتاج الفني - نسخة 0.0.1")
        copyright_label.setAlignment(Qt.AlignCenter)
        copyright_label.setObjectName("CopyrightLabel")
        main_layout.addWidget(copyright_label)
        
        # Header Section
        header_frame = QFrame()
        header_frame.setObjectName("CardFrame")
        header_layout = QHBoxLayout(header_frame)
        header_layout.setContentsMargins(15, 10, 15, 10)
        
        # Logo and Title
        logo_label = QLabel()
        try:
            logo_pixmap = QPixmap()
            try:
                response = requests.get("https://deermedia.co/wp-content/uploads/2020/06/%D8%A7%D9%84%D8%BA%D8%B2%D8%A7%D9%84-.png", timeout=10)
                response.raise_for_status()  # Raise an HTTPError for bad responses
                logo_pixmap.loadFromData(response.content)
            except requests.exceptions.RequestException as e:
                print(f"Error fetching logo image: {e}")
                logo_label = QLabel("DeerMedia")
                logo_label.setFont(QFont("Segoe UI", 16, QFont.Bold))
            logo_label.setPixmap(logo_pixmap.scaledToHeight(60, Qt.SmoothTransformation))
        except:
            logo_label = QLabel("DeerMedia")
            logo_label.setFont(QFont("Segoe UI", 16, QFont.Bold))
        
        title_label = QLabel("WhatsApp Messenger")
        title_label.setObjectName("HeaderTitle")
        title_label.setFont(QFont("Segoe UI", 14, QFont.Bold))
        
        header_layout.addWidget(logo_label)
        header_layout.addSpacing(10)
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        main_layout.addWidget(header_frame)

        # Tab Widget with scroll area
        tab_widget = QTabWidget()
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(tab_widget)

        # ==== Send Messages Tab ====
        send_tab = QWidget()
        send_layout = QVBoxLayout(send_tab)
        send_layout.setContentsMargins(10, 10, 10, 10)
        send_layout.setSpacing(15)

        # File Input Section
        file_card = QFrame()
        file_card.setObjectName("CardFrame")
        file_layout = QVBoxLayout(file_card)
        file_layout.setContentsMargins(15, 15, 15, 15)
        file_layout.setSpacing(10)
        
        file_title = QLabel("DATA SOURCE")
        file_title.setObjectName("SectionTitle")
        file_layout.addWidget(file_title)
        
        # Drop Area
        self.drop_area = FileDropArea()
        self.drop_area.setObjectName("DropArea") 
        self.drop_area.fileDropped.connect(self.set_file_path)  # Connect the signal to the slot
        
        file_layout.addWidget(self.drop_area)
        
        # File Path Display
        self.file_path = QLineEdit()
        self.file_path.setReadOnly(True)
        self.file_path.setPlaceholderText("No file selected")
        file_layout.addWidget(self.file_path)
        
        # Delay Input
        delay_layout = QHBoxLayout()
        delay_layout.addWidget(QLabel("Delay Between Messages (seconds):"))
        self.delay_input = QLineEdit("10")
        self.delay_input.setFixedWidth(80)
        self.delay_input.setValidator(QIntValidator(1, 300))
        delay_layout.addWidget(self.delay_input)
        delay_layout.addStretch()
        file_layout.addLayout(delay_layout)
        
        send_layout.addWidget(file_card)

        # Message Template Section
        msg_card = QFrame()
        msg_card.setObjectName("CardFrame")
        msg_layout = QVBoxLayout(msg_card)
        msg_layout.setContentsMargins(15, 15, 15, 15)
        msg_layout.setSpacing(10)
        
        msg_title = QLabel("MESSAGE TEMPLATE")
        msg_title.setObjectName("SectionTitle")
        msg_layout.addWidget(msg_title)
        
        self.message_template = QTextEdit()
        self.message_template.setPlaceholderText("Enter your message template here...\nUse {name} for customer name")
        self.message_template.setMinimumHeight(120)
        msg_layout.addWidget(self.message_template)
        
        send_layout.addWidget(msg_card)

        # Preview Section
        preview_card = QFrame()
        preview_card.setObjectName("CardFrame")
        preview_layout = QVBoxLayout(preview_card)
        preview_layout.setContentsMargins(15, 15, 15, 15)
        preview_layout.setSpacing(10)
        
        preview_header = QHBoxLayout()
        preview_title = QLabel("DATA PREVIEW")
        preview_title.setObjectName("SectionTitle")
        preview_header.addWidget(preview_title)
        preview_header.addStretch()
        
        self.load_preview_btn = QPushButton("Refresh Preview")
        self.load_preview_btn.setObjectName("SecondaryButton")
        self.load_preview_btn.clicked.connect(self.load_preview)
        preview_header.addWidget(self.load_preview_btn)
        
        preview_layout.addLayout(preview_header)
        
        # Data Table with scroll
        self.data_table = QTableWidget(0, 3)
        self.data_table.setHorizontalHeaderLabels(["Name", "Phone Number", "Message Preview"])
        self.data_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.data_table.setAlternatingRowColors(True)
        self.data_table.verticalHeader().setVisible(False)
        self.data_table.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        
        table_scroll = QScrollArea()
        table_scroll.setWidgetResizable(True)
        table_scroll.setWidget(self.data_table)
        preview_layout.addWidget(table_scroll)
        
        send_layout.addWidget(preview_card)

        # Action Buttons
        action_layout = QHBoxLayout()
        action_layout.setSpacing(15)
        action_layout.addStretch()
        
        self.stop_btn = QPushButton("Stop")
        self.stop_btn.setObjectName("DangerButton")
        self.stop_btn.setFixedSize(120, 35)
        self.stop_btn.setEnabled(False)
        self.stop_btn.clicked.connect(self.stop_sending)
        
        self.start_btn = QPushButton("Start Sending")
        self.start_btn.setObjectName("PrimaryButton")
        self.start_btn.setFixedSize(150, 35)
        self.start_btn.clicked.connect(self.start_sending)
        self.start_btn.setEnabled(False)

  
        action_layout.addWidget(self.stop_btn)
        action_layout.addWidget(self.start_btn)
        send_layout.addLayout(action_layout)

        # ==== Log Tab ====
        log_tab = QWidget()
        log_layout = QVBoxLayout(log_tab)
        log_layout.setContentsMargins(10, 10, 10, 10)
        
        self.log_widget = LogWidget()
        log_layout.addWidget(self.log_widget)

        tab_widget.addTab(send_tab, "Send Messages")
        tab_widget.addTab(log_tab, "Activity Log")
        
        main_layout.addWidget(scroll)  # Add scrollable area

        # Status Bar
        self.status_bar = StatusBar()
        main_layout.addWidget(self.status_bar)

        self.setCentralWidget(main_widget)
    def set_file_path(self, file_path):
        self.file_path.setText(file_path)
        self.status_bar.update_status(f"Excel file selected: {os.path.basename(file_path)}", status_type="info")
        self.log_widget.add_log_entry(f"Excel file selected: {file_path}", "info")
        
        try:
            # قراءة ملف Excel
            df = pd.read_excel(file_path)
            message_col = "Message"  # اسم العمود الافتراضي للرسالة
            
            # إذا كان العمود موجودًا، قم بتحميل الرسالة الأولى في خانة التمبلت
            if message_col in df.columns and not df[message_col].isnull().all():
                first_message = str(df[message_col].iloc[0])
                self.message_template.setPlainText(first_message)
            else:
                self.message_template.setPlainText("")  # إذا لم يكن هناك رسالة، اترك الحقل فارغًا
        except Exception as e:
            self.message_template.setPlainText("")  # في حالة وجود خطأ، اترك الحقل فارغًا
            self.log_widget.add_log_entry(f"Failed to load message template: {str(e)}", "error")
    def load_preview(self):
        excel_file = self.file_path.text()
        if not excel_file:
            QMessageBox.warning(self, "Warning", "Please select an Excel file first")
            return
        
        try:
            # قراءة ملف Excel
            name_col = "Customers Name"
            number_col = "Whatsapp Number"
            message_col = "Message"
            
            df = pd.read_excel(excel_file)
            
            # التحقق من وجود الأعمدة
            for col in [name_col, number_col, message_col]:
                if col not in df.columns:
                    QMessageBox.warning(self, "Warning", f"Column '{col}' not found in Excel file")
                    self.log_widget.add_log_entry(f"Column '{col}' not found in Excel file", "error")
                    return
            
            # عرض المعاينة
            self.data_table.setRowCount(0)
            preview_rows = min(5, len(df))
            self.data_table.setRowCount(preview_rows)
            
            template = self.message_template.toPlainText()
            
            for i in range(preview_rows):
                name = str(df.iloc[i][name_col])
                number = str(df.iloc[i][number_col])
                
                # استخدام التمبلت إذا كان موجودًا
                if template:
                    message = template.replace("{name}", name)
                else:
                    message = str(df.iloc[i][message_col]).replace("{name}", name)
                
                self.data_table.setItem(i, 0, QTableWidgetItem(name))
                self.data_table.setItem(i, 1, QTableWidgetItem(number))
                self.data_table.setItem(i, 2, QTableWidgetItem(message))
            
            status_msg = f"Loaded {len(df)} contacts from Excel file"
            self.status_bar.update_status(status_msg, status_type="success")
            self.log_widget.add_log_entry(status_msg, "success")
            
            # تمكين زر الإرسال بعد عرض المعاينة
            self.start_btn.setEnabled(True)
            
        except Exception as e:
            error_msg = f"Failed to load preview: {str(e)}"
            QMessageBox.critical(self, "Error", error_msg)
            self.log_widget.add_log_entry(error_msg, "error")
    def start_sending(self):
            excel_file = self.file_path.text()
            if not excel_file:
                QMessageBox.warning(self, "Warning", "Please select an Excel file first")
                return
            
            try:
                delay = int(self.delay_input.text())
            except ValueError:
                QMessageBox.warning(self, "Warning", "Please enter a valid delay (seconds)")
                return
            
            # If user provided a template, update the message column
            template = self.message_template.toPlainText()
            if template:
                try:
                    df = pd.read_excel(excel_file)
                    message_col = "Message"  # Hardcoded column name
                    
                    # Create a temporary Excel file with our template
                    temp_file = excel_file.replace(".xlsx", "_temp.xlsx")
                    df[message_col] = template
                    df.to_excel(temp_file, index=False)
                    excel_file = temp_file
                    
                    self.log_widget.add_log_entry(f"Created temporary Excel file with message template", "info")
                except Exception as e:
                    error_msg = f"Failed to apply message template: {str(e)}"
                    QMessageBox.critical(self, "Error", error_msg)
                    self.log_widget.add_log_entry(error_msg, "error")
                    return
            
            # Start the sender thread
            self.sender_thread = WhatsAppSender(excel_file, delay)
            
            self.sender_thread.update_progress.connect(self.update_progress)
            self.sender_thread.finished.connect(self.process_finished)
            self.sender_thread.start()
            
            # Update UI
            self.start_btn.setEnabled(False)
            self.stop_btn.setEnabled(True)
            self.status_bar.update_status("Starting WhatsApp Web...", 0, "info")
            self.log_widget.add_log_entry("Started sending process", "info")
        
    def stop_sending(self):
            if self.sender_thread and self.sender_thread.isRunning():
                self.sender_thread.stop()
                self.status_bar.update_status("Stopping process...", status_type="warning")
                self.log_widget.add_log_entry("Stopping sending process...", "warning")
        
    def update_progress(self, value, message, status_type):
            self.status_bar.update_status(message, value, status_type)
            self.log_widget.add_log_entry(message, status_type)
        
    def process_finished(self, success, message):
            self.status_bar.update_status(message, 100 if success else 0, "success" if success else "error")
            self.log_widget.add_log_entry(message, "success" if success else "error")
            self.start_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
            
            if success:
                QMessageBox.information(self, "Complete", message)
            else:
                QMessageBox.warning(self, "Process Stopped", message)

                self.log_widget.add_log_entry("Process was stopped by user", "warning")

if __name__ == "__main__":
   
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())