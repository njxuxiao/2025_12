import os
import pickle
from pydub import AudioSegment
import whisper
import torch

def generate_cache_for_file(filepath, model, whisper_model_name):
    """
    Generate cache for a single file, logic matches the main program.
    """
    if not os.path.exists(filepath):
        print(f"❌ File not found: {filepath}")
        return

    # Cache filename format must match the main program
    cache_filepath = f"{filepath}.{whisper_model_name}.whisper.cache"
    
    if os.path.exists(cache_filepath):
        print(f"⏭️  [Skipped] Cache already exists: {os.path.basename(cache_filepath)}")
        return

    print(f"🔄 Processing: {os.path.basename(filepath)} ...")
    
    try:
        # 1. Transcribe (This is the time-consuming part)
        # Note: Large models might run out of VRAM on smaller GPUs
        result = model.transcribe(filepath, language="ja")
        
        # 2. Process audio slicing
        # Even if continuous mode doesn't need sliced audio, we save it to maintain
        # compatibility with 'repeat_times > 1' mode (sentence playback).
        audio = AudioSegment.from_file(filepath)
        sentences_data = []
        
        for segment in result["segments"]:
            start_ms = int(segment['start'] * 1000)
            end_ms = int(segment['end'] * 1000)
            text = segment['text'].strip()
            if not text: continue
            
            # Slice the audio segment
            audio_segment = audio[start_ms:end_ms]
            
            sentences_data.append({
                'audio': audio_segment, 
                'text': text,
                'start': start_ms,
                'end': end_ms
            })
        
        # 3. Save to disk
        if sentences_data:
            with open(cache_filepath, 'wb') as f:
                pickle.dump(sentences_data, f)
            print(f"✅ [Success] Generated: {os.path.basename(cache_filepath)}")
        else:
            print(f"⚠️ [Warning] No speech detected: {os.path.basename(filepath)}")

    except Exception as e:
        print(f"❌ [Error] Processing failed: {e}")

def main():
    # ================= Configuration Area =================

    # 1. Set the folder path containing MP3 files
    # The script will automatically scan all .mp3 files in this folder and subfolders
    ROOT_FOLDER = r"./22_12_n3"  # e.g.: r"D:\JapaneseListening\21_7"

    # 2. Or, if you only want to process specific files, list their full paths here
    SPECIFIC_FILES = [
        # r"D:\JapaneseListening\test1.mp3",
        # r"D:\JapaneseListening\test2.mp3",
    ]

    # 3. Set model size (Must match the one used in the main program)
    # Options: "tiny", "base", "small", "medium", "large"
    WHISPER_MODEL = "large"

    # ======================================================

    # 1. Collect all files to process
    files_to_process = []
    
    # Scan folder
    if os.path.exists(ROOT_FOLDER):
        print(f"Scanning folder: {ROOT_FOLDER} ...")
        for root, dirs, files in os.walk(ROOT_FOLDER):
            for file in files:
                if file.lower().endswith('.mp3'):
                    files_to_process.append(os.path.join(root, file))
    
    # Add specific files
    for f in SPECIFIC_FILES:
        if f and os.path.exists(f):
            files_to_process.append(f)
            
    # Remove duplicates
    files_to_process = list(set(files_to_process))
    
    if not files_to_process:
        print("No MP3 files found. Please check the path configuration.")
        return

    print(f"Found {len(files_to_process)} MP3 files.")

    # 2. Load model (Only once)
    device = "cuda" if torch.cuda.is_available() else "cpu"
    print(f"\nLoading Whisper model '{WHISPER_MODEL}' (Device: {device.upper()})...")
    try:
        model = whisper.load_model(WHISPER_MODEL, device=device)
    except Exception as e:
        print(f"Model load failed: {e}")
        return
    print("Model loaded. Starting batch tasks. Time for a coffee! ☕\n")
    print("-" * 50)

    # 3. Loop through files
    for i, filepath in enumerate(files_to_process):
        print(f"[{i+1}/{len(files_to_process)}] ", end="")
        # Pass the model name explicitly to the generator function
        generate_cache_for_file(filepath, model, WHISPER_MODEL)
        
    print("-" * 50)
    print("\n🎉 All tasks completed!")

if __name__ == '__main__':
    main()

    for i in range(10): print("\n")