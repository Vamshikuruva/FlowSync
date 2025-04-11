# 🤖 FlowSync - Your screen’s smartest pixel.

A smart floating assistant for your desktop that:
- Detects open documents (PDFs, DOCX, Excel, etc.)
- Captures screen content using OCR
- Accepts queries through a floating UI
- Answers questions using LLMs (OpenAI / Gemini)
- Automates tasks on your screen using Python + keyboard control
- Supports multiple document uploads

---

## ✨ Features

- 🖼️ **Floating Chat UI** (hotkey-activated)
- 📄 **Smart document analysis** (Word, PDF, Excel, PowerPoint, etc.)
- 🧠 **OpenAI-powered Q&A** from documents
- 🔍 **Screen OCR + assistant suggestions**
- ⚙️ **Keyboard automation scripting**
- 📁 Supports **multiple document uploads**
- ⌨️ **Ctrl+Alt+A** to activate

---

## 📦 Installation

```bash
git clone https://github.com/your-username/ai-desktop-assistant.git
cd ai-desktop-assistant

# (Optional) Create virtual environment
python -m venv venv
source venv/bin/activate  # or venv\Scripts\activate

# Install dependencies
pip install -r requirements.txt
```

---

## 🔑 Environment Variables

Create a `.env` file in the project root with your API keys:

```
OPENAI_API_KEY=your-openai-key
GOOGLE_API_KEY=your-gemini-key
```

---

## 🚀 Usage

```bash
python hotkey_launcher.py
```

Then, press **Ctrl+Alt+A** to toggle the assistant. You can:

- Ask about the open document
- Add additional documents
- Ask screen-related queries
- Trigger automation

---

## 📂 Supported Document Formats

- `.docx`
- `.pdf`
- `.txt`
- `.xlsx`, `.xls`
- `.pptx`, `.ppt`
- `.csv`
- `.json`

---

## 🧠 Tech Stack

- **LangChain** for chaining LLM + embeddings
- **OpenAI GPT-3.5** for Q&A
- **Gemini / GPT** for screen analysis + code generation
- **FAISS** for document embeddings
- **EasyOCR** for screen reading
- **PyQt5** for UI
- **PyAutoGUI + Keyboard** for automation

---

## 🎯 Hotkey

> Press **Ctrl+Alt+A** to toggle the floating chat assistant UI.

---

## 🧪 Use Cases
Summarize or answer questions about a research paper open in Word

Get suggestions or automate actions based on what’s on-screen

Perform tasks like launching apps, typing messages, filling forms

Switch between document expert mode and screen assistant mode instantly

---

## 🛠️ Todo / Improvements

- [ ] Add dark/light theme toggle
- [ ] Drag-to-move collapsed UI
- [ ] Display uploaded document names
- [ ] Option to clear document history

---

## 📄 License

MIT License

