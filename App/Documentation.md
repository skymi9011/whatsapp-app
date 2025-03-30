<div dir="rtl">

### **Documentation: شرح أجزاء التطبيق**


---

### **1. المكتبات المستوردة (Lines 1-20)**

```python
import sys
import os
import pandas as pd
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from urllib.parse import quote
from PyQt5.QtWidgets import ...
from PyQt5.QtGui import ...
from PyQt5.QtCore import ...
import requests
```

#### **شرح:**
- **`sys` و `os`**: لإدارة النظام ومسارات الملفات.
- **`pandas`**: مكتبة لمعالجة ملفات Excel.
- **`time`**: لإضافة تأخيرات بين الرسائل.
- **`selenium`**: للتحكم في متصفح Chrome وفتح WhatsApp Web.
- **`PyQt5`**: لإنشاء واجهة المستخدم الرسومية (GUI).
- **`requests`**: لتحميل الصور أو البيانات من الإنترنت.

---

### **2. تعريف الألوان (Lines 22-30)**

```python
PRIMARY_COLOR = "#1976D2"  # Deeper blue
SECONDARY_COLOR = "#00BFA5"  # Teal
BG_COLOR = "#F8F9FA"  # Lighter background
TEXT_COLOR = "#2D3748"  # Darker text for better contrast
SUCCESS_COLOR = "#4CAF50"  # Success green
WARNING_COLOR = "#FF9800"  # Warning orange
ERROR_COLOR = "#F44336"  # Error red
CARD_SHADOW = "rgba(0, 0, 0, 0.1)"  # Shadow color for cards
```

#### **شرح:**
- هذه الألوان تُستخدم لتنسيق واجهة المستخدم (مثل الأزرار، النصوص، والخلفيات).
- كل لون له غرض محدد:
  - **`PRIMARY_COLOR`**: اللون الأساسي.
  - **`SUCCESS_COLOR`**: لون النجاح.
  - **`ERROR_COLOR`**: لون الخطأ.

---

### **3. كلاس `WhatsAppSender` (Lines 32-141)**

#### **أ. تعريف الكلاس (Lines 32-41)**

```python
class WhatsAppSender(QThread):
    update_progress = pyqtSignal(int, str, str)  # إشارة لتحديث التقدم
    finished = pyqtSignal(bool, str)  # إشارة عند انتهاء العملية
```

- **`QThread`**: يسمح بتشغيل العملية في خلفية التطبيق دون تجميد واجهة المستخدم.
- **`update_progress`**: تُستخدم لتحديث شريط التقدم.
- **`finished`**: تُستخدم لإعلام التطبيق بانتهاء العملية.

---

#### **ب. تهيئة الكلاس (Lines 42-49)**

```python
def __init__(self, excel_file, delay, parent=None):
    super().__init__(parent)
    self.excel_file = excel_file
    self.delay = delay
    self.message_column = "Message"
    self.name_column = "Customers Name"
    self.number_column = "Whatsapp Number"
    self.is_running = True
```

- **`excel_file`**: مسار ملف Excel الذي يحتوي على بيانات العملاء.
- **`delay`**: التأخير بين الرسائل (بالثواني).
- **`message_column`**: اسم العمود الذي يحتوي على الرسائل.
- **`name_column`**: اسم العمود الذي يحتوي على أسماء العملاء.
- **`number_column`**: اسم العمود الذي يحتوي على أرقام الواتساب.

---

#### **ج. إيقاف العملية (Lines 50-51)**

```python
def stop(self):
    self.is_running = False
```

- تُستخدم لإيقاف العملية إذا أراد المستخدم إيقاف الإرسال.

---

#### **د. تشغيل العملية (Lines 52-141)**

##### **1. قراءة ملف Excel (Lines 54-57)**

```python
df = pd.read_excel(self.excel_file)
if df[self.number_column].dtype != 'object':
    df[self.number_column] = df[self.number_column].astype(str)
df[self.number_column] = df[self.number_column].str.replace(r'\D+', '', regex=True)  # تنظيف الأرقام
```

- يقرأ ملف Excel ويتأكد من أن الأرقام مكتوبة كنصوص.
- يُزيل أي رموز أو مسافات من الأرقام.

---

##### **2. فتح متصفح Chrome (Lines 59-63)**

```python
chrome_profile_path = os.path.join(os.environ['USERPROFILE'], 'AppData', 'Local', 'Google', 'Chrome', 'User Data', 'Default')
options = webdriver.ChromeOptions()
options.add_argument(f'user-data-dir={chrome_profile_path}')
driver = webdriver.Chrome(options=options)
driver.get('https://web.whatsapp.com/')
```

- يفتح متصفح Chrome باستخدام Selenium.
- يستخدم ملف تعريف Chrome لتجنب تسجيل الدخول في كل مرة.

---

##### **3. التحقق من تسجيل الدخول (Lines 65-77)**

```python
try:
    WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.XPATH, '//div[@role="textbox"][@contenteditable="true"]'))
    )
    self.update_progress.emit(0, "Logged into WhatsApp Web successfully.", "success")
except Exception:
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
```

- يتحقق إذا كان المستخدم مسجل الدخول بالفعل.
- إذا لم يكن مسجل الدخول، يطلب منه مسح الباركود (QR Code).

---


### **4. كلاس `StyleSheet` (Lines 141-410)**

#### **أ. تعريف الكلاس (Lines 141-143)**

```python
class StyleSheet:
    @staticmethod
    def get_stylesheet():
        return f"""
        ...
        """
```

- **`StyleSheet`**: كلاس يحتوي على التنسيقات (CSS) الخاصة بواجهة المستخدم.
- **`get_stylesheet`**: دالة ثابتة تُرجع النصوص التنسيقية (CSS) لتطبيقها على عناصر واجهة المستخدم.

---

#### **ب. التنسيقات الأساسية (Base Styling) (Lines 145-150)**

```css
QMainWindow, QWidget {
    background-color: #F8FAFC;
    color: #1E293B;
    font-family: 'Segoe UI', -apple-system, system-ui;
    font-size: 13px;
}
```

- **`QMainWindow` و `QWidget`**: يتم تطبيق هذه التنسيقات على النافذة الرئيسية وجميع عناصر الواجهة.
- **`background-color`**: لون الخلفية.
- **`font-family`**: نوع الخط المستخدم.
- **`font-size`**: حجم الخط الافتراضي.

---

#### **ج. تنسيقات الأزرار (Buttons) (Lines 170-190)**

```css
QPushButton {
    border: 1px solid #CBD5E1;
    border-radius: 6px;
    padding: 8px 16px;
    background: white;
    color: #334155;
    font-weight: 500;
    min-width: 100px;
}
```

- **`QPushButton`**: تنسيق الأزرار الافتراضية.
- **`border`**: حدود الزر.
- **`border-radius`**: زوايا مستديرة.
- **`background`**: لون الخلفية.
- **`color`**: لون النص.

---

#### **د. تنسيقات شريط التقدم (Progress Bar) (Lines 210-220)**

```css
QProgressBar {
    border: 1px solid #E2E8F0;
    border-radius: 4px;
    height: 18px;
    text-align: center;
}
QProgressBar::chunk {
    background: #2563EB;
    border-radius: 3px;
}
```

- **`QProgressBar`**: تنسيق شريط التقدم.
- **`QProgressBar::chunk`**: تنسيق الجزء الذي يمثل التقدم.

---

### **5. كلاس `LogWidget` (Lines 412-440)**

#### **أ. تعريف الكلاس (Lines 412-414)**

```python
class LogWidget(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        ...
```

- **`LogWidget`**: ويدجت (Widget) مخصص لعرض السجلات (Logs) أثناء تشغيل التطبيق.

---

#### **ب. إضافة السجلات (Lines 424-432)**

```python
def add_log_entry(self, message, status_type="info"):
    color_map = {
        "success": SUCCESS_COLOR,
        "warning": WARNING_COLOR,
        "error": ERROR_COLOR,
        "info": PRIMARY_COLOR
    }
    ...
```

- **`add_log_entry`**: تضيف رسالة جديدة إلى السجل.
- **`status_type`**: نوع الرسالة (نجاح، تحذير، خطأ، معلومات).

---

### **6. كلاس `StatusBar` (Lines 442-480)**

#### **أ. تعريف الكلاس (Lines 442-444)**

```python
class StatusBar(QWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        ...
```

- **`StatusBar`**: شريط الحالة الذي يعرض الرسائل وشريط التقدم.

---

#### **ب. تحديث شريط الحالة (Lines 460-478)**

```python
def update_status(self, message, progress=None, status_type="info"):
    icon_map = {
        "success": "✓",
        "warning": "⚠",
        "error": "✗",
        "info": "ℹ"
    }
    ...
```

- **`update_status`**: تُحدث الرسالة وشريط التقدم بناءً على الحالة الحالية.
- **`status_type`**: نوع الحالة (نجاح، تحذير، خطأ، معلومات).

---

### **7. كلاس `FileDropArea` (Lines 482-540)**

#### **أ. تعريف الكلاس (Lines 482-484)**

```python
class FileDropArea(QFrame):
    fileDropped = pyqtSignal(str)
    ...
```

- **`FileDropArea`**: ويدجت مخصص للسماح للمستخدم بسحب وإفلات ملفات Excel.

---

#### **ب. التعامل مع السحب والإفلات (Lines 500-510)**

```python
def dropEvent(self, event):
    urls = event.mimeData().urls()
    if urls and urls[0].toLocalFile():
        file_path = urls[0].toLocalFile()
        if file_path.endswith(('.xlsx', '.xls')):
            self.fileDropped.emit(file_path)
        else:
            QMessageBox.warning(self, "Invalid File", "Please drop a valid Excel file.")
```

- **`dropEvent`**: يتحقق من الملف المسحوب ويقبله إذا كان ملف Excel.

---

### **8. كلاس `MainWindow` (Lines 542-1010)**

#### **أ. تعريف الكلاس (Lines 542-544)**

```python
class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.initUI()
```

- **`MainWindow`**: النافذة الرئيسية للتطبيق.

---

#### **ب. واجهة المستخدم (Lines 546-1010)**

##### **1. إعداد الواجهة (Lines 546-550)**

```python
def initUI(self):
    self.setStyleSheet(StyleSheet.get_stylesheet())
    self.setWindowTitle("DeerMedia WhatsApp Messenger")
    ...
```

- **`initUI`**: تُنشئ جميع عناصر واجهة المستخدم.

##### **2. تحميل ملف Excel (Lines 950-960)**

```python
def set_file_path(self, file_path):
    self.file_path.setText(file_path)
    ...
```

- **`set_file_path`**: تُحدث مسار الملف عند اختياره.

##### **3. بدء الإرسال (Lines 970-990)**

```python
def start_sending(self):
    excel_file = self.file_path.text()
    ...
```

- **`start_sending`**: تبدأ عملية الإرسال باستخدام البيانات الموجودة في ملف Excel.

---

### **9. تشغيل التطبيق (Lines 1012-1016)**

```python
if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
```

- **`__main__`**: نقطة البداية لتشغيل التطبيق.

---
