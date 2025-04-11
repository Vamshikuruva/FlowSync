import os
import re
import sys
import time
import tempfile
import shutil
import json
import keyboard
import easyocr
import pyautogui
import pygetwindow as gw
import pyperclip
import subprocess
import logging
from dotenv import load_dotenv
from PIL import Image
from langchain_core.prompts import ChatPromptTemplate
from langchain.schema.output_parser import StrOutputParser
from langchain_google_genai import ChatGoogleGenerativeAI
from langchain_openai import ChatOpenAI

# === Load Environment ===
load_dotenv()
google_api_key = os.getenv("GOOGLE_API_KEY")
openai_api_key = os.getenv("OPENAI_API_KEY")

# === Logging ===
logging.basicConfig(filename="assistant.log", level=logging.INFO, format="%(asctime)s - %(message)s")

# === Temp Directory for Screenshots ===
TEMP_DIR = tempfile.gettempdir()
SCREENSHOT_PATH = os.path.join(TEMP_DIR, "screen_capture.png")

# === OCR ===
ocr_reader = easyocr.Reader(['en'], gpu=False)

# === Assistant Modes ===
ASSISTANT_MODE = "smart"  # "fast" or "smart"
MAX_HISTORY = 5
conversation_history = []

def format_conversation_history(history):
    formatted = ""
    for item in history[-MAX_HISTORY:]:
        if item.get("type") == "instruction":
            formatted += f"\n---\nUser Query: {item['query']}\nGemini Instructions: {item['instructions'][:100]}...\n"
        elif item.get("type") == "automation_attempt":
            formatted += f"\n---\nAutomation Attempt:\nCode: {item['code_attempt'][:100]}...\nError: {item['error']}\n"
        elif item.get("type") == "automation_success":
            formatted += f"\n---\nAutomation Success:\nCode: {item['code_attempt'][:100]}...\n"
        else:
            formatted += f"\n---\n[Unrecognized history entry]\n{json.dumps(item, indent=2)}\n"
    return formatted

# === LangChain Gemini Setup ===
# llm_general = ChatGoogleGenerativeAI(model="gemini-1.5-pro-latest", temperature=0.3, api_key=google_api_key)
# llm_code = ChatGoogleGenerativeAI(model="gemini-1.5-pro-latest", temperature=0.2, api_key=google_api_key)
llm_general = ChatOpenAI(model="gpt-3.5-turbo", temperature=0.3, api_key=openai_api_key)
llm_code = ChatOpenAI(model="gpt-3.5-turbo", temperature=0.2, api_key=openai_api_key)

# === Prompts ===
suggestion_prompt = ChatPromptTemplate.from_template("""
You are an intelligent assistant. The following is a snapshot of the user's screen:

---SCREEN CONTENT---
{screen}
---------------------

Give a list of tasks the user might want help with based solely on this screen.
Be concise and specific.
""")

query_prompt = ChatPromptTemplate.from_template("""
You are an expert desktop assistant helping a user on a Windows 11 system.

The user sees the following on their screen:
"{screen}"

The user just asked:
"{query}"

Only use the conversation history below if it contains an unfinished or directly related query:
{history}

Your task:
1. First, give a simple, beginner-friendly set of manual instructions that do **not** use keyboard shortcuts. These should be very clear and helpful even for someone unfamiliar with computers.

2. Then write a **Python automation script** that performs this task using **only keyboard-based automation**. Use only:
   - `keyboard.press_and_release()`
   - `keyboard.write()`
   - `pyautogui.hotkey()`
   - `pyautogui.press()`
   - `pyperclip.copy()` + `pyautogui.hotkey("ctrl", "v")`

⚠️ AUTOMATION RULES:
- Always add `time.sleep(1)` **after each key press or action** to ensure stable execution.
- Avoid using `keyboard.press()` and `keyboard.release()` manually unless absolutely required. Instead, use `keyboard.press_and_release()` or `pyautogui.hotkey()`.
- Do not use `pyautogui.click()` or move the mouse unless there's absolutely no other keyboard-based alternative.
- Use `time.sleep(8-10)` when opening an app, and `time.sleep(2-5)` when loading content or navigating.
- At the **end of your script**, include safe fallback code to release all keys if any were held (e.g., using `keyboard.release('ctrl')`, `keyboard.release('alt')`, etc.).
- Your automation should never leave keys pressed down.

Respond **ONLY** in the following JSON format:
```json
{{
  "instructions": "Step-by-step manual instructions go here.",
  "automation_code": "Python code using keyboard-only shortcuts with 1s delays and key release safety(if applicable)."
}}
""")



# === Intent Detection ===
def classify_query_intent(user_query):
    intent_prompt = ChatPromptTemplate.from_template("""
    Classify the user's query into one of two categories:
    - "automation" if the user is asking to perform a desktop action or automate something.
    - "general" if the user is asking a question, explanation, or non-automatable information.

    User Query: "{query}"
    Category:
    """)
    chain = intent_prompt | llm_general | StrOutputParser()
    response = chain.invoke({"query": user_query}).strip().lower()
    return "automation" if "automation" in response else "general"

# === Core Chain Handlers ===
def suggest_task_from_screen(screen_content):
    chain = suggestion_prompt | llm_general | StrOutputParser()
    return chain.invoke({"screen": screen_content})

def respond_to_user_query(screen_content, user_query):
    use_history = ASSISTANT_MODE == "smart"
    history_formatted = format_conversation_history(conversation_history) if use_history else ""

    query_type = classify_query_intent(user_query)

    if query_type == "general":
        # Just respond as a chatbot
        print("💬 Detected general query. Responding conversationally.")
        chain = ChatPromptTemplate.from_template("""
        You are a smart and helpful assistant. Your job is to answer the user's question.

        Context from screen (if any): 
        "{screen}"

        User Question: 
        "{query}"

        Respond clearly and helpfully. If the screen content is not relevant, ignore it.
        Use your own general knowledge or reasoning. Only refer to screen content if it's necessary.
        """) | llm_general | StrOutputParser()

        return chain.invoke({"screen": screen_content, "query": user_query}), ""

    # Automation case
    chain = query_prompt | llm_code | StrOutputParser()
    result = chain.invoke({
        "screen": screen_content,
        "query": user_query,
        "history": history_formatted
    })
    instructions, code = parse_response(result)

    if use_history:
        conversation_history.append({
            "screen_context": screen_content,
            "query": user_query,
            "instructions": instructions,
            "automation_code": code,
            "type": "instruction"
        })
    return instructions, code


# === response parsing ===

def parse_response(response):
    try:
        # Step 1: Extract JSON object (whether inside a code block or raw)
        code_block_match = re.search(r"\{[\s\S]*\}", response)
        if not code_block_match:
            raise ValueError("No JSON found in the response.")

        json_str = code_block_match.group(0).strip()

        # Step 2: Parse JSON
        parsed = json.loads(json_str)

        # Step 3: Normalize instructions (list or string)
        instructions_raw = parsed.get("instructions", "")
        if isinstance(instructions_raw, list):
            instructions = "\n".join(instructions_raw)
        else:
            instructions = instructions_raw.strip()

        # Step 4: Return both
        automation_code = parsed.get("automation_code", "").strip()
        return instructions, automation_code

    except Exception as e:
        print("❌ Failed to parse ChatGPT response:", e)
        return response.strip(), ""

# === Automation Execution ===
def execute_code(code_str, screen_context, user_query, max_attempts=3):
    print("💻 Executing automation code...")
    attempts = 0
    current_code = code_str.strip().replace("```python", "").replace("```", "")
    while attempts < max_attempts:
        try:
            # keyboard.clear_all_hotkeys()  # Clears held keys from hotkey listener
            time.sleep(2)  # Give time for the user to prepare
            exec(current_code, globals())
            print("✅ Task automated successfully.")
            conversation_history.append({
                "screen": screen_context,
                "query": user_query,
                "code_attempt": current_code,
                "status": "success",
                "type": "automation_success"
            })
            return
        except Exception as e:
            print(f"❌ Automation failed (Attempt {attempts+1}/{max_attempts}): {e}")
            logging.error(f"Automation failed on attempt {attempts+1}: {e}")
            conversation_history.append({
                "screen": screen_context,
                "query": user_query,
                "code_attempt": current_code,
                "error": str(e),
                "status": "failed",
                "type": "automation_attempt"
            })
            fix_prompt = ChatPromptTemplate.from_template("""
                You are an expert Python developer helping to fix broken automation code.
                The following Python automation script was generated to help the user complete a task based on their desktop screen.
                However, it failed with an error. Your job is to fix ONLY the error while keeping the original logic and structure intact.
                📄 Screen Context: {screen}
                ❓ User Query: {query}

                💻 Original Python Code:
                ```python
                {current_code}
                ❗ Error Message: "" {e} ""
                ✅ Your task:
                Fix the above Python code so it works correctly.
                Return only the corrected Python code (no markdown formatting, no explanations, no comments).
                DO NOT wrap the code in a function.
                Keep the code as close to the original as possible. Only fix what is broken.
                Please output only raw working Python code. Nothing else.""")
            # Format the fix_prompt template first
            formatted_prompt = fix_prompt.format(
                screen=screen_context,
                query=user_query,
                current_code=current_code,
                e=str(e)
            )
            # Now invoke the LLM with the formatted prompt (as a string)
            fixed_code = llm_code.invoke(formatted_prompt)
            print("\n🔧 Gemini suggests a fixed version:\n")
            print(fixed_code)
            confirm = input("\n⚙️ Do you want to try the fixed code? (y/n): ").strip().lower()
            if confirm != "y":
                print("⛔ Skipping automation.")
                break
            current_code = fixed_code.strip().replace("```python", "").replace("```", "")
            attempts += 1
    print("❌ Could not complete automation after multiple attempts.")
    return

# === Highlighting Click Locations ===
def highlight_and_click(text_to_find):
    try:
        result = ocr_reader.readtext(SCREENSHOT_PATH, detail=1)
        for (bbox, text, _) in result:
            if text_to_find.lower() in text.lower():
                (top_left, top_right, bottom_right, bottom_left) = bbox
                x = int((top_left[0] + bottom_right[0]) / 2)
                y = int((top_left[1] + bottom_right[1]) / 2)
                pyautogui.moveTo(x, y, duration=0.3)
                pyautogui.click()
                pyautogui.sleep(0.5)
                print(f"✅ Clicked on '{text}'")
                return True
        print("⚠️ No matching element found to click.")
        return False
    except Exception as e:
        print(f"❌ Error in highlight_and_click: {e}")
        logging.error(f"highlight_and_click error: {e}")
        return False

# === Helpers ===
def contains_code(response):
    return any(cmd in response for cmd in ["pyautogui", "pyperclip", "subprocess", "webbrowser", "keyboard", "time"])

def capture_and_process_screen(region=None):
    if os.path.exists(SCREENSHOT_PATH):
        os.remove(SCREENSHOT_PATH)
    image = pyautogui.screenshot(region=region)
    image.save(SCREENSHOT_PATH)
    result = ocr_reader.readtext(SCREENSHOT_PATH, detail=0)
    return "\n".join(result).strip()

# === Background Listener ===
def start_background_listener():
    global ASSISTANT_MODE
    print(f"📣 Assistant running in {ASSISTANT_MODE.upper()} mode... Press Ctrl+L for capturing screen and Press ESC anytime to exit.")
    while True:
        if keyboard.is_pressed('ctrl+l'):
            print("\n🟠 Capturing screen...")
            extracted_text = capture_and_process_screen()
            if not extracted_text:
                print("⚠️ No text detected on screen.")
                continue

            suggestion = suggest_task_from_screen(extracted_text)
            print(f"\n💡 Gemini Suggests: {suggestion}")

            while True:
                user_query = input("\n❓ What do you want help with? (type 'exit' to recapture, 'mode' to switch): ").strip()
                if user_query.lower() == "exit":
                    print(f"\n📣 Assistant running in {ASSISTANT_MODE.upper()} mode... Press Ctrl+L for capturing screen and Press ESC anytime to exit.")
                    break
                elif user_query.lower() == "mode":
                    user_mode = input("Enter any mode (smart/fast): ").strip().lower()
                    if user_mode not in ["smart", "fast"]:
                        print("⚠️ Invalid mode. Please enter 'smart' or 'fast'.")
                        continue
                    ASSISTANT_MODE = user_mode
                    if ASSISTANT_MODE == "fast":
                        conversation_history.clear()
                    print(f"🔁 Switched to {ASSISTANT_MODE.upper()} mode.")
                    continue

                instructions, automation_code = respond_to_user_query(extracted_text, user_query)
                print(f"\n💡 Gemini Response:\n\n{instructions}")

                if contains_code(automation_code.strip()):
                    should_do = input("\n⚙️ Should I perform this task? (y/n): ").strip().lower()
                    if should_do == "y":
                        execute_code(automation_code.strip(), extracted_text, user_query)
                else:
                    print("\nℹ️ This was a general response. No automation will be performed.")

        elif keyboard.is_pressed('esc'):
            print("\n👋 Exiting assistant.")
            break
        time.sleep(0.2)

# === Entry Point ===
if __name__ == "__main__":
    start_background_listener()
