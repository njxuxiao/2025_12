import os
import sys
import re
import shutil  # Added to dynamically get terminal width

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

def display_box(title, text, width=80, border_style='double'):
    """Draw a console text box"""
    if border_style == 'double':
        top_l, top_r, bot_l, bot_r, h_line, v_line = "╔", "╗", "╚", "╝", "═", "║"
    else:
        top_l, top_r, bot_l, bot_r, h_line, v_line = "╭", "╮", "╰", "╯", "─", "│"

    print("\n" + top_l + h_line * width + top_r)
    
    title_line = f" [{title}] "
    title_pad = width - get_display_length(title_line)
    # Ensure title padding doesn't go negative
    title_pad = max(0, title_pad)
    print(v_line + title_line + " " * title_pad + v_line)
    print(v_line + " " * width + v_line)

    lines = wrap_text(text, width - 4)
    for line in lines:
        padding = width - 4 - get_display_length(line)
        padding = max(0, padding)
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
            elif key_stroke.lower() == 'j': return 'j'
            elif key_stroke in ['\r', '\n']: return 'enter' 
    else:
        while True:
            event = keyboard.read_event(suppress=True)
            if event.event_type == keyboard.KEY_DOWN:
                name = event.name.lower()
                if name == 'right': return 'right'
                elif name == 'left': return 'left'
                elif name in ['q', 'j', 'enter']: return name

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

        idx = 0

        while idx < len(sentences):
            os.system('cls' if os.name == 'nt' else 'clear')
            current_sentence = sentences[idx]

            print("--- Immersive Japanese Reader ---")
            
            progress_title = f"Original Text [{idx + 1}/{len(sentences)}]"
            
            # ==== Dynamic Width Calculation ====
            # Get actual terminal window width (default to 120 if undetectable)
            term_width = shutil.get_terminal_size((120, 24)).columns
            
            # Calculate how much width the sentence actually needs (+8 for padding and borders)
            needed_width = get_display_length(current_sentence) + 8
            
            # Set minimum width to 80, but expand if the sentence is longer
            dynamic_width = max(80, needed_width)
            
            # Cap the maximum width to fit strictly inside the terminal window to prevent ugly line breaks
            final_width = min(dynamic_width, term_width - 2)

            # Display original Japanese text with Progress Tracking
            display_box(progress_title, current_sentence, width=final_width, border_style='double')

            # Operation prompt
            print("\n" + "="*70)
            print(f"Controls: [→] Next | [←] Prev | [J] Jump")
            print(f"          [Enter] Boss Key (Clear Screen) | [Q] Quit")
            
            key = wait_for_key()

            if key == 'right':
                idx += 1
            elif key == 'left':
                idx = max(0, idx - 1)
            elif key == 'j':
                # ==== Jump Logic ====
                print("\n[Jump Mode]", end="")
                try:
                    jump_input = input(f" Enter sentence number (1-{len(sentences)}): ")
                    if jump_input.strip(): 
                        jump_idx = int(jump_input.strip()) - 1
                        if 0 <= jump_idx < len(sentences):
                            idx = jump_idx
                except ValueError:
                    pass
            elif key == 'enter':
                # ==== Boss Key (Anti-peeping) Logic ====
                os.system('cls' if os.name == 'nt' else 'clear')
                fake_prompt = f"{os.getcwd()}>" if os.name == 'nt' else f"{os.environ.get('USER', 'user')}@localhost:~$"
                print(fake_prompt, end="", flush=True)
                wait_for_key()
            elif key == 'q':
                print("\nReading ended.")
                break

    except Exception as e:
        print(f"Program error: {e}")

if __name__ == '__main__':
    # Replace the txt file path here with your Japanese article file path
    article_file = r"D:\N2_2025_12\QUORA\1.txt"
    read_article(article_file)