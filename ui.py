import sys
import os
os.environ["KMP_DUPLICATE_LIB_OK"] = "TRUE"

from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, QTextEdit,
                             QLineEdit, QDesktopWidget, QHBoxLayout, QMessageBox, QSystemTrayIcon, QStyle, QFileDialog)
from PyQt5.QtCore import Qt, QRectF, QTimer
from PyQt5.QtGui import QRegion, QPainterPath, QColor
from PyQt5.QtWidgets import QGraphicsDropShadowEffect
from screen import capture_and_process_screen, suggest_task_from_screen, respond_to_user_query, execute_code
from detect_open import detect_document_path, extract_text, build_temp_index_from_file, close_application_by_pid, copy_to_temp, reopen_file, FILE_TYPES
from dotenv import load_dotenv
from langchain.chains import RetrievalQA
from langchain_openai import ChatOpenAI

load_dotenv()
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY")

class FloatingChat(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowFlags(Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint | Qt.Tool)
        self.setAttribute(Qt.WA_TranslucentBackground)

        self.circle_radius = 40
        self.expanded = False

        self.screen_text = ""
        self.document_text = ""
        self.document_index = None
        self.document_mode = False

        self.resize(self.circle_radius * 2, self.circle_radius * 2)
        self.move_to_bottom_right()
        self.init_ui()
        self.set_circle_shape()

    def move_to_bottom_right(self):
        screen_rect = QDesktopWidget().availableGeometry()
        x = screen_rect.width() - self.width() - 20
        y = screen_rect.height() - self.height() - 20
        self.move(x, y)

    def set_circle_shape(self):
        # Delay applying the shape slightly to ensure correct geometry is available
        QTimer.singleShot(0, self._apply_circle_mask)

    def _apply_circle_mask(self):
        path = QPainterPath()
        rectf = QRectF(self.rect())
        path.addEllipse(rectf)
        region = QRegion(path.toFillPolygon().toPolygon())
        self.setMask(region)

    def init_ui(self):
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

        self.button = QPushButton("💬")
        self.button.setFixedSize(self.circle_radius * 2, self.circle_radius * 2)
        self.button.clicked.connect(self.toggle_expand)

        # Add glow effect
        glow = QGraphicsDropShadowEffect(self)
        glow.setBlurRadius(20)
        glow.setColor(QColor("#00ffff"))  # Cyan glow
        glow.setOffset(0, 0)
        self.button.setGraphicsEffect(glow)

        # Add pulsing animation (opacity)
        from PyQt5.QtCore import QPropertyAnimation, QEasingCurve
        self.pulse_anim = QPropertyAnimation(self.button, b"windowOpacity")
        self.pulse_anim.setDuration(1500)
        self.pulse_anim.setStartValue(1.0)
        self.pulse_anim.setEndValue(0.7)
        self.pulse_anim.setEasingCurve(QEasingCurve.InOutQuad)
        self.pulse_anim.setLoopCount(-1)
        self.pulse_anim.start()
        self.layout.addWidget(self.button)

        self.chat_box = QTextEdit()
        self.chat_box.setPlaceholderText("Assistant responses will appear here...")
        self.chat_box.setReadOnly(True)
        self.chat_box.setVisible(False)
        self.chat_box.setMinimumHeight(160)

        self.input_field = QLineEdit()
        self.input_field.setPlaceholderText("Type your question and press Enter...")
        self.input_field.returnPressed.connect(self.handle_user_query)
        self.input_field.setVisible(False)

        self.mode_toggle_btn = QPushButton()
        self.mode_toggle_btn.setVisible(False)
        self.mode_toggle_btn.clicked.connect(self.toggle_mode)

        self.add_doc_btn = QPushButton("➕ Add Document")
        self.add_doc_btn.setVisible(False)
        self.add_doc_btn.clicked.connect(self.add_new_document)
        self.update_mode_button()

        self.close_button = QPushButton("❌ Close")
        self.close_button.clicked.connect(self.close_app)
        self.close_button.setVisible(False)

        self.layout.addWidget(self.chat_box)
        self.layout.addWidget(self.mode_toggle_btn)
        self.layout.addWidget(self.add_doc_btn)
        self.layout.addWidget(self.input_field)
        self.layout.addWidget(self.close_button)

        self.tray_icon = QSystemTrayIcon(self)
        self.tray_icon.setIcon(self.style().standardIcon(QStyle.SP_ComputerIcon))
        self.tray_icon.show()

    def toggle_expand(self):
        if self.expanded:
            # Minimize to circle
            self.chat_box.setVisible(False)
            self.input_field.setVisible(False)
            self.mode_toggle_btn.setVisible(False)
            self.add_doc_btn.setVisible(False)
            self.close_button.setVisible(False)
            self.resize(self.circle_radius * 2, self.circle_radius * 2)
            self.set_circle_shape()
            self.move_to_bottom_right()
            self.button.setVisible(True)  # Fixed: use self.button instead of self.chat_button
            self.expanded = False
        else:
            # Expand to assistant mode
            self.setMask(QRegion())  # Clear circular mask
            self.resize(420, 400)
            self.move_to_bottom_right()
            self.chat_box.setVisible(True)
            self.input_field.setVisible(True)
            self.mode_toggle_btn.setVisible(True)
            self.add_doc_btn.setVisible(self.document_mode)
            self.close_button.setVisible(True)
            self.button.setVisible(False)  # Fixed: use self.button instead of self.chat_button
            self.expanded = True
            self.initialize_context()

    def show_toast(self, message):
        self.tray_icon.showMessage("Assistant", message, QSystemTrayIcon.Information, 3000)

    def update_mode_button(self):
        if self.document_mode:
            self.mode_toggle_btn.setText("🧠 Mode: Document Expert → Click to switch")
        else:
            self.mode_toggle_btn.setText("🖥️ Mode: Screen Assistant → Click to switch")
        self.add_doc_btn.setVisible(self.expanded and self.document_mode)

    def toggle_mode(self):
        self.document_mode = not self.document_mode
        self.update_mode_button()
        if self.document_mode:
            self.chat_box.append("\n🧠 Switched to Document Expert mode.")
        else:
            self.chat_box.append("\n🖥️ Switched to Screen Assistant mode.")

    def initialize_context(self):
        self.chat_box.setText("🔍 Checking your screen and open documents...")
        self.show_toast("Analyzing screen and checking for open documents...")
        QApplication.processEvents()
        try:
            file_path, process = detect_document_path()
            process_name = process.name().lower() if process else ""
            ext = os.path.splitext(file_path)[-1].lower() if file_path else ""

            print("[DEBUG] file_path:", file_path)
            print("[DEBUG] process:", process_name)

            if file_path and ext in FILE_TYPES:
                close_application_by_pid(process.pid)
                temp_path = copy_to_temp(file_path)
                self.document_text = extract_text(temp_path)
                self.document_index = build_temp_index_from_file(temp_path, OPENAI_API_KEY)
                reopen_file(file_path)

                self.chat_box.setText(
                    f"📄 A supported document is open: {os.path.basename(file_path)}\n\nYou're now in Document Expert mode. Ask your question below."
                )
                self.document_mode = True
                self.mode_toggle_btn.setVisible(True)
                self.update_mode_button()
                self.layout.update()
                self.updateGeometry()
                return

            elif process_name in ["winword.exe", "excel.exe", "powerpnt.exe"]:
                self.chat_box.setText(
                    f"📄 A document app is open (e.g., Word/Excel/PowerPoint), but file access failed.\n\nDefaulting to Document Expert mode."
                )
                self.document_mode = True
                self.mode_toggle_btn.setVisible(True)
                self.update_mode_button()
                self.layout.update()
                self.updateGeometry()
                return

            # fallback: only screen
            self.screen_text = capture_and_process_screen()
            screen_suggestions = suggest_task_from_screen(self.screen_text)
            self.chat_box.setText(f"💡 Gemini Suggestions (Screen):\n{screen_suggestions}\n\nAsk anything below.")
            self.show_toast("No document detected. Using screen context.")
            self.document_mode = False
            self.mode_toggle_btn.setVisible(True)
            self.update_mode_button()

        except Exception as e:
            self.chat_box.setText(f"❌ Error during initialization: {e}")
            self.show_toast("Initialization error. See assistant for details.")

    def add_new_document(self):
        file_path, _ = QFileDialog.getOpenFileName(self, "Select a document", "", "Documents (*.pdf *.docx *.txt)")
        if file_path:
            try:
                self.chat_box.append(f"\n📄 Loading: {os.path.basename(file_path)}")
                temp_path = copy_to_temp(file_path)
                self.document_text = extract_text(temp_path)
                self.document_index = build_temp_index_from_file(temp_path, OPENAI_API_KEY)
                self.chat_box.append("✅ Document added and ready to query.")
            except Exception as e:
                self.chat_box.append(f"❌ Failed to load document: {e}")

    def handle_user_query(self):
        user_query = self.input_field.text().strip()
        if not user_query:
            return
        self.input_field.clear()
        self.chat_box.append(f"\n🧑 You: {user_query}\n")
        QApplication.processEvents()

        try:
            if self.document_mode and self.document_index:
                retriever = self.document_index.as_retriever()
                qa_chain = RetrievalQA.from_chain_type(
                    llm=ChatOpenAI(model="gpt-3.5-turbo", temperature=0.2, api_key=OPENAI_API_KEY),
                    retriever=retriever,
                    return_source_documents=False
                )
                answer = qa_chain.invoke(user_query)
                self.chat_box.append(f"🤖 Document Answer:\n{answer['result']}")
            else:
                instructions, automation_code = respond_to_user_query(self.screen_text, user_query)
                self.chat_box.append(f"🤖 Screen Assistant:\n{instructions}\n")
                if automation_code.strip():
                    confirm = QMessageBox.question(
                        self, "Automation Request",
                        "Do you want me to perform this task automatically?",
                        QMessageBox.Yes | QMessageBox.No
                    )
                    if confirm == QMessageBox.Yes:
                        self.chat_box.append("⚙️ Running automation...\n")
                        QApplication.processEvents()
                        result = execute_code(automation_code, self.screen_text, user_query)
                        self.chat_box.append(f"✅ Automated successfully.\n{result if result else ''}")
        except Exception as e:
            self.chat_box.append(f"❌ Error: {e}")

    def close_app(self):
        QApplication.quit()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    chat = FloatingChat()
    chat.show()
    sys.exit(app.exec_())
