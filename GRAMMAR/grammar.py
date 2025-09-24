import pandas as pd
import os
import sys
import tempfile
import time

try:
    import keyboard
except ImportError:
    print("Error: 'keyboard' library is missing.")
    print("Please run in your terminal or command line: pip install keyboard")
    sys.exit()

try:
    import pykakasi
except ImportError:
    print("Error: 'pykakasi' library is missing.")
    print("Please run in your terminal or command line: pip install pykakasi")
    sys.exit()

try:
    from gtts import gTTS
    from playsound import playsound
    gtts_available = True
except ImportError:
    print("Warning: gTTS or playsound library not installed. Online TTS function will be unavailable.")
    print("  - pip install gTTS")
    print("  - pip install playsound==1.2.2")
    gtts_available = False

try:
    import pyttsx3
    pyttsx3_available = True
except ImportError:
    print("Warning: pyttsx3 library not installed. Offline backup TTS function will be unavailable.")
    print("  - pip install pyttsx3")
    pyttsx3_available = False


def get_pyttsx3_japanese_voice_id():
    if not pyttsx3_available:
        return None
    try:
        temp_engine = pyttsx3.init()
        voices = temp_engine.getProperty('voices')
        temp_engine.stop()
        del temp_engine

        japanese_voice_id = None
        for voice in voices:
            lang_str = getattr(voice, 'lang', '').lower()
            name_str = getattr(voice, 'name', '').lower()
            if 'ja' in lang_str or 'japanese' in name_str:
                japanese_voice_id = voice.id
                break
        
        if japanese_voice_id:
            print("\nSuccessfully found [Local] Japanese voice ID. Offline backup is available.")
            return japanese_voice_id
        else:
            print("\nWarning: Japanese [Local] voice package not found in your system. Offline backup may not work.")
            return None
    except Exception as e:
        print(f"\nFailed to scan for local voices: {e}")
        return None


def speak_with_gtts(text):
    if not gtts_available or not text:
        return False
    
    temp_mp3 = None
    fp = None
    success = False

    try:
        fp = tempfile.NamedTemporaryFile(delete=False, suffix='.mp3')
        temp_mp3 = fp.name
        fp.close()
        fp = None

        for attempt in range(3):
            try:
                tts = gTTS(text=text, lang='ja')
                tts.save(temp_mp3)
                success = True
                break 
            except Exception as e:
                print(f"\n!! Online TTS connection failed (Attempt {attempt + 1}/3): {e}")
                if attempt < 2:
                    time.sleep(0.5)

        if success:
            playsound(temp_mp3)
        return success
    except Exception as e:
        print(f"\n!! An unknown error occurred during speech synthesis: {e}")
        return False
    finally:
        if fp is not None:
            fp.close()
        if temp_mp3 and os.path.exists(temp_mp3):
            try:
                os.remove(temp_mp3)
            except Exception as e:
                print(f"!! Failed to clean up temporary file: {e}")


def speak_with_pyttsx3(voice_id, text):
    if voice_id and text:
        try:
            engine = pyttsx3.init()
            engine.setProperty('voice', voice_id)
            engine.say(text)
            engine.runAndWait()
            engine.stop()
        except Exception as e:
            print(f"\n!! Error using local TTS: {e}")


def get_display_length(s):
    length = 0
    for char in str(s):
        if '\u4e00' <= char <= '\u9fff':
            length += 2
        else:
            length += 1
    return length

def display_term(word, grammar):
    print("\n" + "╔" + "═"*50 + "╗")
    print("║" + " "*50 + "║")
    
    if word:
        line = f"  [Word]: {word}"
        padding = 50 - get_display_length(line)
        if padding < 0: padding = 0
        print("║" + line + " "*padding + "║")
        
    if grammar:
        line = f"  [Grammar]: {grammar}"
        padding = 50 - get_display_length(line)
        if padding < 0: padding = 0
        print("║" + line + " "*padding + "║")

    print("║" + " "*50 + "║")
    print("╚" + "═"*50 + "╝")

def display_details(meaning, remarks):
    if not meaning and not remarks:
        return

    print("╭" + "┈"*50 + "╮")
    
    if meaning:
        line = f"  [Meaning]: {meaning}"
        padding = 50 - get_display_length(line)
        if padding < 0: padding = 0
        print("┆" + line + " "*padding + "┆")

    if remarks:
        line = f"  [Remarks]: {remarks}"
        padding = 50 - get_display_length(line)
        if padding < 0: padding = 0
        print("┆" + line + " "*padding + "┆")

    print("╰" + "┈"*50 + "╯")


def study_helper(file_path, tts_mode='auto'):
    japanese_voice_id = get_pyttsx3_japanese_voice_id()
    auto_mode_current_engine = 'gTTS' if gtts_available else 'pyttsx3'

    kks = pykakasi.kakasi()
    df = None

    try:
        if not os.path.exists(file_path):
            print(f"Error: File not found '{file_path}'.")
            return
        
        if not file_path.lower().endswith('.txt'):
            print(f"Error: This script only supports .txt files. Please provide a tab-delimited text file.")
            return

        print(f"Reading from TXT file: {os.path.basename(file_path)}")
        df = pd.read_csv(file_path, sep='\t')
        df.columns = df.columns.str.strip()

    except Exception as e:
        print(f"Error reading TXT file: {e}")
        return

    if 'Fre' not in df.columns:
        print("No 'Fre' column detected, creating it automatically.")
        df['Fre'] = 0
    else:
        df['Fre'] = pd.to_numeric(df['Fre'], errors='coerce').fillna(0).astype(int)

    if '单词' not in df.columns and '文法' not in df.columns:
        print(f"Error: The data must contain at least a '单词' or '文法' column.")
        return
        
    has_reading_col = '读音' in df.columns
    has_meaning_col = '含义' in df.columns
    has_remarks_col = '备注' in df.columns
    
    df.sort_values(by='Fre', ascending=False, inplace=True)
    
    print("\n--- Japanese Study Helper (TXT Mode) Started ---")
    if tts_mode == 'online':
        print("[TTS Mode]: Online Only")
    elif tts_mode == 'offline':
        print("[TTS Mode]: Offline Only")
    else:
        print("[TTS Mode]: Auto (Online first, fallback to Offline)")
        
    print("[IMPORTANT] Please make sure your input method is in English mode to ensure key presses are registered correctly.")
    
    is_changed = False
    last_answered_correctly_index = None
    records = df.to_dict('records')
    
    i = 0
    enter_press_count = 0
    while i < len(records):
        current_record = records[i]

        word = str(current_record.get('单词', '')) if pd.notna(current_record.get('单词')) else ""
        grammar = str(current_record.get('文法', '')) if pd.notna(current_record.get('文法')) else ""
        
        reading = ""
        if has_reading_col and pd.notna(current_record.get('读音')):
            reading = str(current_record.get('读音'))
        elif word:
            try:
                result = kks.convert(word)
                reading = "".join([item['hira'] for item in result])
            except Exception as e:
                print(f"!! Could not convert '{word}' to hiragana: {e}")

        meaning = str(current_record.get('含义', '')) if has_meaning_col and pd.notna(current_record.get('含义')) else ""
        remarks = str(current_record.get('备注', '')) if has_remarks_col and pd.notna(current_record.get('备注')) else ""

        if not word and not grammar:
            i += 1
            continue
        
        is_cleared = False
        key = None
        while True:
            if not is_cleared:
                display_term(word, grammar)

            prompt = "Press a key... (→: Know / 0: Don't Know / q: Quit"
            if last_answered_correctly_index is not None:
                prompt += " / x: Correct Last"
            prompt += " / Enter: New Line)"
            print(prompt + " " + str(i+1) + "/" + str(len(records)))
            
            event = keyboard.read_event(suppress=True)
            while event.event_type != keyboard.KEY_DOWN:
                event = keyboard.read_event(suppress=True)
            
            key = event.name.lower()

            if key == 'enter':
                enter_press_count += 1
                if enter_press_count > 3:
                    os.system('cls' if os.name == 'nt' else 'clear')
                    is_cleared = True
                else:
                    print("\n" * 2)
                    is_cleared = False
                continue
            
            enter_press_count = 0
            break
        
        if key == 'q':
            print("Saving progress and exiting...")
            break 
        
        if key == 'x':
            if last_answered_correctly_index is not None:
                records[last_answered_correctly_index]['Fre'] += 1
                is_changed = True
                print(f"\nCorrected the previous item!")
                last_answered_correctly_index = None
            else:
                print("\nThere is no previous item to correct.")
            continue

        last_answered_correctly_index = None
        text_to_speak = word if word else grammar
        display_remarks = remarks if remarks else reading

        def speak():
            if tts_mode == 'online':
                if not speak_with_gtts(text_to_speak): print("\n!! Online TTS failed.")
            elif tts_mode == 'offline':
                if japanese_voice_id: speak_with_pyttsx3(japanese_voice_id, text_to_speak)
                else: print("\n!! Offline TTS is not available.")
            else: # Auto mode
                nonlocal auto_mode_current_engine
                if auto_mode_current_engine == 'gTTS':
                    if not speak_with_gtts(text_to_speak):
                        print("\n!! Online TTS failed, switching to [Offline TTS] mode.")
                        auto_mode_current_engine = 'pyttsx3'
                        speak_with_pyttsx3(japanese_voice_id, text_to_speak)
                else:
                    speak_with_pyttsx3(japanese_voice_id, text_to_speak)

        if key == '0':
            current_record['Fre'] += 1
            is_changed = True
            print(f"Recorded! Forgotten count: {current_record['Fre']}")
            display_details(meaning, display_remarks)
            speak()

        elif key == 'right':
            display_details(meaning, display_remarks)
            speak()
            print(f"Great! Forgotten count: {current_record['Fre']}")
            last_answered_correctly_index = i

        i += 1

    print("\nAll words/grammar have been studied!")

    if not is_changed:
        print("\nNo changes were made, no need to save.")
        return
        
    try:
        print("Saving progress to TXT file...")
        new_df = pd.DataFrame(records)
        new_df.to_csv(file_path, sep='\t', index=False, encoding='utf-8-sig')
        print("\nStudy session finished! Progress saved to TXT.")
    except PermissionError:
        print(f"\nError saving file: Permission denied. Please close '{file_path}' and try again.")
    except Exception as e:
        print(f"\nAn unknown error occurred while saving the file: {e}")


if __name__ == '__main__':
    
    # --- HOW TO USE (At Work) ---
    # 1. Place your converted .txt files in the same folder as this script.
    # 2. Add the names of your .txt files to the list below.
    txt_list_work = [
        "1_21_07_Sheet1.txt", 
        "2_22_12_Sheet1.txt", 
        "3_22_07_Sheet1.txt",
        "4_21_12_Sheet1.txt",
        "5_20_12_Sheet1.txt",
        "6_19_12_Sheet1.txt",
        "7_19_07_Sheet1.txt"
    ]

    # 3. Choose which file to study by its index (0 is the first one).
    file_to_study_index = 0
    
    if file_to_study_index >= len(txt_list_work):
        print(f"Error: Index {file_to_study_index} is out of range for the file list.")
        sys.exit()

    selected_file_path = txt_list_work[file_to_study_index]
    
    # 4. Set your preferred TTS engine ('auto', 'online', or 'offline')
    preferred_tts_engine = 'offline' 

    study_helper(selected_file_path, tts_mode=preferred_tts_engine)
