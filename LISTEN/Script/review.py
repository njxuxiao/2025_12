import pandas as pd
import os
import sys
from openpyxl import load_workbook
import tempfile
import time

try:
    import keyboard
except ImportError:
    print("Error: 'keyboard' library is missing.")
    print("Please run in your terminal or command line: pip install keyboard")
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


def study_helper(file_path, sheet_to_study=None, tts_mode='auto'):
    japanese_voice_id = get_pyttsx3_japanese_voice_id()
    auto_mode_current_engine = 'gTTS' if gtts_available else 'pyttsx3'

    all_sheets_data = {}
    chosen_sheet = None
    try:
        if not os.path.exists(file_path):
            print(f"Error: File not found '{file_path}'.")
            return

        if not file_path.endswith('.xlsx'):
            print(f"Error: This feature requires an .xlsx format Excel file.")
            print(f"Please open '{os.path.basename(file_path)}' with Excel and save it as .xlsx format.")
            return
        
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        sheet_names = xls.sheet_names

        if not sheet_names:
            print("Error: No worksheets found in the Excel file.")
            return

        if sheet_to_study is not None:
            if isinstance(sheet_to_study, int):
                if 0 <= sheet_to_study < len(sheet_names):
                    chosen_sheet = sheet_names[sheet_to_study]
                else:
                    print(f"Error: Specified sheet index {sheet_to_study} is invalid. Valid range is 0 to {len(sheet_names)-1}.")
                    return
            elif isinstance(sheet_to_study, str):
                if sheet_to_study in sheet_names:
                    chosen_sheet = sheet_to_study
                else:
                    print(f"Error: Cannot find worksheet named '{sheet_to_study}'.")
                    print(f"Available worksheets are: {sheet_names}")
                    return
            else:
                print("Error: Invalid type for sheet_to_study parameter. It should be an integer (index) or a string (name).")
                return
            print(f"Selected worksheet as specified: '{chosen_sheet}'")
        else:
            if len(sheet_names) == 1:
                chosen_sheet = sheet_names[0]
                print(f"Automatically selected the only worksheet: '{chosen_sheet}'")
            else:
                print("Multiple worksheets (Sheets) found:")
                for i, name in enumerate(sheet_names):
                    print(f"  {i+1}: {name}")
                while True:
                    try:
                        choice = int(input(f"Please enter the number of the worksheet you want to study (1-{len(sheet_names)}): "))
                        if 1 <= choice <= len(sheet_names):
                            chosen_sheet = sheet_names[choice-1]
                            break
                        else:
                            print("Invalid number, please try again.")
                    except ValueError:
                        print("Please enter a number.")
        
        all_sheets_data = pd.read_excel(file_path, sheet_name=None, engine='openpyxl')
        df = all_sheets_data[chosen_sheet]

    except Exception as e:
        print(f"Error reading or selecting worksheet: {e}")
        return

    if 'Fre' not in df.columns:
        print("No 'Fre' column detected, creating it automatically.")
        df['Fre'] = 0
    else:
        df['Fre'] = pd.to_numeric(df['Fre'], errors='coerce').fillna(0).astype(int)

    if '单词' not in df.columns and '文法' not in df.columns:
        print(f"Error: Worksheet '{chosen_sheet}' must contain at least a '单词' or '文法' column.")
        return
        
    has_meaning_col = '含义' in df.columns
    has_remarks_col = '备注' in df.columns
    if not has_meaning_col:
        print("Info: No '含义' (Meaning) column in your Excel, definitions will not be shown.")
    if not has_remarks_col:
        print("Info: No '备注' (Remarks) column in your Excel, remarks will not be shown.")

    df.sort_values(by='Fre', ascending=False, inplace=True)
    print("\nSorted by 'Fre' (Frequency). The most forgotten items will appear first.")

    print("\n--- Japanese Study Helper Started ---")
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
    original_indices = df.index.tolist()

    i = 0
    while i < len(records):
        current_record = records[i]
        original_index = original_indices[i]

        word = str(current_record.get('单词', '')) if pd.notna(current_record.get('单词')) else ""
        grammar = str(current_record.get('文法', '')) if pd.notna(current_record.get('文法')) else ""
        meaning = str(current_record.get('含义', '')) if has_meaning_col and pd.notna(current_record.get('含义')) else ""
        remarks = str(current_record.get('备注', '')) if has_remarks_col and pd.notna(current_record.get('备注')) else ""

        if not word and not grammar:
            i += 1
            continue
        
        display_term(word, grammar)

        prompt = "Press a key... (→: Know / 0: Don't Know / q: Quit"
        if last_answered_correctly_index is not None:
            prompt += " / x: Correct Last)"
        else:
            prompt += ")"
        print(prompt + " " + str(i) + "/" + str(len(records)))
        
        event = keyboard.read_event(suppress=True)
        while event.event_type != keyboard.KEY_DOWN:
            event = keyboard.read_event(suppress=True)
        
        key = event.name.lower()
        
        if key == 'q':
            print("Saving progress and exiting...")
            break 
        
        if key == 'x':
            if last_answered_correctly_index is not None:
                for idx, record in enumerate(records):
                    if original_indices[idx] == last_answered_correctly_index:
                        record['Fre'] += 1
                        break
                is_changed = True
                print(f"\nCorrected the previous item!")
                last_answered_correctly_index = None
            else:
                print("\nThere is no previous item to correct.")
            continue

        last_answered_correctly_index = None
        text_to_speak = word if word else grammar

        def speak():
            if tts_mode == 'online':
                if not speak_with_gtts(text_to_speak):
                    print("\n!! Online TTS failed.")
                return

            if tts_mode == 'offline':
                if japanese_voice_id:
                    speak_with_pyttsx3(japanese_voice_id, text_to_speak)
                else:
                    print("\n!! Offline TTS is not available.")
                return
            
            nonlocal auto_mode_current_engine
            if auto_mode_current_engine == 'gTTS':
                if not speak_with_gtts(text_to_speak):
                    print("\n!! Online TTS failed, automatically switching to [Offline TTS] mode.")
                    auto_mode_current_engine = 'pyttsx3'
                    speak_with_pyttsx3(japanese_voice_id, text_to_speak)
            elif auto_mode_current_engine == 'pyttsx3':
                speak_with_pyttsx3(japanese_voice_id, text_to_speak)

        if key == '0':
            current_record['Fre'] += 1
            is_changed = True
            print(f"Recorded! Forgotten count: {current_record['Fre']}")
            display_details(meaning, remarks)
            speak()

        elif key == 'right':
            display_details(meaning, remarks)
            speak()
            print("Great! print("Great! Forgotten count: {current_record['Fre']}")")
            last_answered_correctly_index = original_index

        i += 1

    print("\nAll words/grammar have been studied!")

    if not is_changed:
        print("\nNo changes were made, no need to save.")
        return
        
    try:
        print("Updating records back to DataFrame...")
        new_df = pd.DataFrame(records, index=original_indices)
        all_sheets_data[chosen_sheet] = new_df

        print("Saving file and preserving column widths...")
        book = load_workbook(file_path)
        col_widths = {}
        for sheet_name in book.sheetnames:
            col_widths[sheet_name] = {
                letter: dim.width for letter, dim in book[sheet_name].column_dimensions.items()
            }
        
        with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
            for sheet_name, sheet_data in all_sheets_data.items():
                sheet_data.to_excel(writer, sheet_name=sheet_name, index=True)
                
                if sheet_name in col_widths:
                    ws = writer.sheets[sheet_name]
                    for col_letter, width in col_widths[sheet_name].items():
                        if width:
                           ws.column_dimensions[col_letter].width = width

        print("\nStudy session finished! Your progress has been saved successfully, and column widths are preserved.")
    except PermissionError:
        print(f"\nError saving file: Permission denied. Please close the Excel file '{file_path}' and try again.")
    except Exception as e:
        print(f"\nAn unknown error occurred while saving the file: {e}")


if __name__ == '__main__':
    excel_file_path = './21_7/21_7.xlsx' 
    
    study_sheet = 0

    preferred_tts_engine = 'auto' 

    study_helper(excel_file_path, sheet_to_study=study_sheet, tts_mode=preferred_tts_engine)

