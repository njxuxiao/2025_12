import os
import sys
import time
import re

# ================= Dependency Check =================
try:
    if os.name == 'nt':
        import msvcrt
except ImportError:
    pass

try:
    import keyboard
except ImportError:
    print("Error: 'keyboard' library is missing. (pip install keyboard)")
    sys.exit()

try:
    import argostranslate.package
    import argostranslate.translate
except ImportError:
    print("Error: 'argostranslate' library is missing.")
    print("Please run in terminal: pip install argostranslate")
    sys.exit()

# ================= Offline Translation Model Initialization =================
def ensure_offline_translation_models():
    """Check and download language packages for offline translation (requires internet only for the first run)"""
    try:
        installed = argostranslate.package.get_installed_packages()
        installed_codes = [(pkg.from_code, pkg.to_code) for pkg in installed]
        
        # argostranslate translates via English as an intermediate language
        needed_paths = [('ja', 'en'), ('en', 'zh')]
        missing = [p for p in needed_paths if p not in installed_codes]
        
        if missing:
            print("First time using offline translation, downloading local language packages (requires brief internet connection, may take a few minutes)...")
            argostranslate.package.update_package_index()
            available = argostranslate.package.get_available_packages()
            
            for p in missing:
                for pkg in available:
                    if pkg.from_code == p[0] and pkg.to_code == p[1]:
                        print(f"Downloading {p[0]} -> {p[1]} language package...")
                        argostranslate.package.install_from_path(pkg.download())
                        break
            print("Offline translation model configuration complete! Future translations will run completely offline.\n")
            time.sleep(2)
    except Exception as e:
        print(f"Offline language package initialization failed: {e}")
        print("Please check your network and try again, or ensure you have enough disk space.")
        sys.exit()

# ================= UI & Formatting Module =================
def get_display_length(s):
    """Calculate the terminal display length of a string (full-width characters count as 2, half-width as 1)"""
    length = 0
    for char in str(s):
        if ('\u4e00' <= char <= '\u9fff' or 
            '\u3040' <= char <= '\u30ff' or 
            '\uff00' <= char <= '\uffef'):
            length += 2
        else:
            length += 1
    return length

def wrap_text(text, max_width):
    """Automatically wrap long sentences based on terminal width"""
    lines = []
    current_line = ""
    current_len = 0
    for char in text:
        char_len = 2 if ('\u4e00' <= char <= '\u9fff' or '\u3040' <= char <= '\u30ff' or '\uff00' <= char <= '\uffef') else 1
        if current_len + char_len > max_width:
            lines.append(current_line)
            current_line = char
            current_len = char_len
        else:
            current_line += char
            current_len += char_len
    if current_line:
        lines.append(current_line)
    return lines

def display_box(title, text, width=60, border_style='double'):
    """Draw a console text box"""
    if border_style == 'double':
        top_l, top_r, bot_l, bot_r, h_line, v_line = "╔", "╗", "╚", "╝", "═", "║"
    else:
        top_l, top_r, bot_l, bot_r, h_line, v_line = "╭", "╮", "╰", "╯", "─", "│"

    print("\n" + top_l + h_line * width + top_r)
    
    title_line = f" [{title}] "
    title_pad = width - get_display_length(title_line)
    print(v_line + title_line + " " * title_pad + v_line)
    print(v_line + " " * width + v_line)

    lines = wrap_text(text, width - 4)
    for line in lines:
        padding = width - 4 - get_display_length(line)
        print(v_line + "  " + line + " " * padding + "  " + v_line)
        
    print(v_line + " " * width + v_line)
    print(bot_l + h_line * width + bot_r)

# ================= Core Reading Logic =================
def split_into_sentences(text):
    """Split the article into sentences based on common Japanese punctuation marks and line breaks"""
    raw_sentences = re.split(r'(?<=[。！？\n])', text)
    sentences = [s.strip() for s in raw_sentences if s.strip()]
    return sentences

def wait_for_key():
    """Listen for keyboard input"""
    if os.name == 'nt':
        while True:
            key_stroke = msvcrt.getwch()
            if key_stroke in ['\xe0', '\x00']:
                key_stroke = msvcrt.getwch()
                if key_stroke == 'M': return 'right'
                elif key_stroke == 'K': return 'left'
            elif key_stroke.lower() == 'q': return 'q'
            elif key_stroke.lower() == 't': return 't'
            elif key_stroke in ['\r', '\n']: return 'enter' # Boss Key triggers on Enter
    else:
        while True:
            event = keyboard.read_event(suppress=True)
            if event.event_type == keyboard.KEY_DOWN:
                name = event.name.lower()
                if name == 'right': return 'right'
                elif name == 'left': return 'left'
                elif name in ['q', 't', 'enter']: return name

def read_article(file_path):
    try:
        if not os.path.exists(file_path):
            print(f"Error: File not found '{file_path}'")
            return

        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()

        sentences = split_into_sentences(content)
        if not sentences:
            print("The article is empty or no valid sentences were recognized.")
            return

        # Check offline models
        ensure_offline_translation_models()

        idx = 0
        show_translation = False
        translation_cache = {}

        while idx < len(sentences):
            os.system('cls' if os.name == 'nt' else 'clear')
            current_sentence = sentences[idx]

            print("--- Immersive Japanese Reader ---")
            
            # Display original Japanese text with Progress Tracking
            progress_title = f"Original Text [{idx + 1}/{len(sentences)}]"
            display_box(progress_title, current_sentence, width=70, border_style='double')

            # Display translation
            if show_translation:
                print(" Translating offline...\r", end="")
                if idx not in translation_cache:
                    try:
                        # Call offline translation engine
                        translation_cache[idx] = argostranslate.translate.translate(current_sentence, 'ja', 'zh')
                    except Exception:
                        translation_cache[idx] = f"[Translation Failed: Offline translation engine error or model not loaded correctly]"
                
                print(" " * 25 + "\r", end="") # Clear 'translating' prompt
                display_box("Chinese Translation", translation_cache[idx], width=70, border_style='single')

            # # Operation prompt
            # print("\n" + "="*70)
            # print(f"Controls: [→] Next | [←] Prev | [T] Translate Current")
            # print(f"          [Enter] Boss Key (Clear Screen) | [Q] Quit")
            
            key = wait_for_key()

            if key == 'right':
                idx += 1
                show_translation = False
            elif key == 'left':
                idx = max(0, idx - 1)
                show_translation = False
            elif key == 't':
                show_translation = True
            elif key == 'enter':
                # ==== Boss Key (Anti-peeping) Logic ====
                os.system('cls' if os.name == 'nt' else 'clear')
                # Fake a minimalist system command line interface
                fake_prompt = f"{os.getcwd()}>" if os.name == 'nt' else f"{os.environ.get('USER', 'user')}@localhost:~$"
                print(fake_prompt, end="", flush=True)
                # Wait for any valid key (right, left, t, q, or enter) to restore
                wait_for_key()
                # Hide translation upon restoration for better anti-peeping
                show_translation = False 
            elif key == 'q':
                print("\nReading ended.")
                break

    except Exception as e:
        print(f"Program error: {e}")

if __name__ == '__main__':
    # Replace the txt file path here with your Japanese article file path
    article_file = "1.txt"
    read_article(article_file)
    for i in range(10):
        print("\n" * 5)