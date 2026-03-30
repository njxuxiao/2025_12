import os
import threading
import queue
import pickle
import io
from pydub import AudioSegment
import time
import whisper
import numpy as np
import librosa
import pygame
from pynput import keyboard
import torch

def play_audio(segment):
    buffer = io.BytesIO()
    segment.export(buffer, format="wav")
    buffer.seek(0)
    sound = pygame.mixer.Sound(buffer)
    sound.play()

def apply_speed_change(segment, speed=1.0):
    if speed == 1.0:
        return segment

    dtype_map = {1: np.int8, 2: np.int16, 4: np.int32}
    if segment.sample_width not in dtype_map:
        print(f"Warning: Unsupported audio bit width ({segment.sample_width}), falling back to low-quality speed change.")
        return segment._spawn(segment.raw_data, overrides={"frame_rate": int(segment.frame_rate * speed)})

    dtype = dtype_map[segment.sample_width]
    samples = np.frombuffer(segment.raw_data, dtype=dtype)
    
    samples_float = samples.astype(np.float32) / np.iinfo(dtype).max
    
    if segment.channels == 2:
        samples_float = samples_float.reshape((-1, 2)).T

    stretched_samples = librosa.effects.time_stretch(y=samples_float, rate=speed)

    if segment.channels == 2:
        stretched_samples = stretched_samples.T.flatten()

    y = (stretched_samples * np.iinfo(dtype).max).astype(dtype)

    return AudioSegment(
        y.tobytes(),
        frame_rate=segment.frame_rate,
        sample_width=segment.sample_width,
        channels=segment.channels
    )

def input_collector(q):
    def on_press(key):
        command = None
        try:
            if key.char in ['r', 'q', 'a', 's', 'd']:
                command = {
                    'r': 'r', 
                    'q': 'q',
                    'a': 'speed_down', 
                    's': 'speed_reset', 
                    'd': 'speed_up'
                }[key.char]
        except AttributeError:
            # Modified: Arrows now map to generic directional commands
            if key == keyboard.Key.right: command = 'right'
            elif key == keyboard.Key.left: command = 'left'
            elif key == keyboard.Key.space: command = 'toggle_pause'
        if command:
            q.put(command)

    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()

def run_sentence_playback_mode(sentences_data, repeat_times, command_queue):
    """
    Original logic for repeat_times > 1: Plays sliced audio segments.
    Arrows (right/left) map to Next/Previous sentence here.
    """
    current_sentence_index = 0
    is_paused = False
    command_to_process = None
    playback_speed = 1.0

    while 0 <= current_sentence_index < len(sentences_data):
        if command_to_process:
            if command_to_process == 'right': # Arrow Right -> Next Sentence
                if current_sentence_index < len(sentences_data) - 1: current_sentence_index += 1
            elif command_to_process == 'left': # Arrow Left -> Prev Sentence
                if current_sentence_index > 0: current_sentence_index -= 1
            elif command_to_process == 'q':
                print("Practice finished. Keep up the good work!")
                break
            command_to_process = None

        sentence_info = sentences_data[current_sentence_index]
        sentence_audio = sentence_info['audio']
        sentence_text = sentence_info['text']
        
        print("\n" + "="*50)
        print(f"--- Sentence {current_sentence_index + 1}/{len(sentences_data)} ---")
        print(f"  Text: {sentence_text}")
        print("="*50)

        playback_interrupted = False
        for i in range(repeat_times):
            segment_with_speed = apply_speed_change(sentence_audio, playback_speed)
            print(f"Playing: {i + 1}/{repeat_times} (Speed: {playback_speed:.1f}x)")
            play_audio(segment_with_speed)
            
            while pygame.mixer.get_busy() or is_paused:
                try:
                    command = command_queue.get_nowait()
                    if command == 'toggle_pause':
                        if is_paused: pygame.mixer.unpause(); is_paused = False; print("[ Resumed ]", end="", flush=True)
                        else: pygame.mixer.pause(); is_paused = True; print("\n[ Paused ]", end="", flush=True)
                    elif command in ['right', 'left', 'q', 'r']:
                        pygame.mixer.stop(); is_paused = False; command_to_process = command; playback_interrupted = True; print(f"\nCommand received, interrupting playback."); break
                    elif command in ['speed_up', 'speed_down', 'speed_reset']:
                        pygame.mixer.stop(); is_paused = False
                        if command == 'speed_up': playback_speed = round(playback_speed + 0.1, 1)
                        elif command == 'speed_down': playback_speed = max(0.5, round(playback_speed - 0.1, 1))
                        elif command == 'speed_reset': playback_speed = 1.0
                        print(f"\n[ Speed changed to: {playback_speed:.1f}x ]"); command_to_process = 'r'; playback_interrupted = True; break
                except queue.Empty:
                    time.sleep(0.1)

            if playback_interrupted: break
            if i < repeat_times - 1: time.sleep(0.4)
        
        if not playback_interrupted:
            if current_sentence_index < len(sentences_data) - 1:
                print("\n[ Auto-playing next sentence... ]")
                time.sleep(0.5)
                command_to_process = 'right' # Auto-next
            else:
                print("\n*** Reached end. Controls: (←) Prev, (r) Repeat, (q) Quit, (a/s/d) Speed ***")
                command = command_queue.get()
                if command in ['speed_up', 'speed_down', 'speed_reset']:
                    if command == 'speed_up': playback_speed = round(playback_speed + 0.1, 1)
                    elif command == 'speed_down': playback_speed = max(0.5, round(playback_speed - 0.1, 1))
                    elif command == 'speed_reset': playback_speed = 1.0
                    print(f"\n[ Speed set to: {playback_speed:.1f}x ]"); command_to_process = 'r'
                else:
                    command_to_process = command

def run_continuous_mode(filepath, sentences_data, command_queue):
    """
    New logic for repeat_times == 1: Plays file directly.
    Arrows (right/left) map to Seek +5s/-5s here.
    """
    print("\nStarting Continuous Mode (Direct Playback).")
    print("Controls: (Space) Pause, (→) +5s, (←) -5s, (d/a/s) Speed, (q) Quit")
    
    # Initialize Audio
    # We load the file directly into pygame mixer for streaming (fast startup)
    pygame.mixer.music.load(filepath)
    pygame.mixer.music.play()
    
    playback_speed = 1.0
    is_paused = False
    
    # To track position accurately across seeks and speed changes
    # start_time_offset: where in the audio file (seconds) we started playing
    current_playback_start_offset = 0.0
    
    # We keep a copy in RAM for speed processing if needed
    full_audio_segment = None 

    current_sentence_index = -1
    
    # Helper to load data for speed processing only when needed
    def get_full_audio():
        nonlocal full_audio_segment
        if full_audio_segment is None:
            print("Loading full audio into RAM for processing (this happens once)...")
            full_audio_segment = AudioSegment.from_file(filepath)
        return full_audio_segment

    def play_at_position(pos_seconds):
        nonlocal current_playback_start_offset
        # Ensure pos is within bounds
        pos_seconds = max(0.0, pos_seconds)
        # We can't check max duration easily without pydub, but pygame handles EOF
        
        pygame.mixer.music.play(start=pos_seconds)
        current_playback_start_offset = pos_seconds
        if is_paused:
            pygame.mixer.music.pause()

    while True:
        try:
            command = command_queue.get_nowait()
            
            # --- Quit ---
            if command == 'q':
                pygame.mixer.music.stop()
                print("\nPractice finished.")
                break
            
            # --- Pause ---
            if command == 'toggle_pause':
                if is_paused:
                    pygame.mixer.music.unpause()
                    is_paused = False
                    print("[ Resumed ]", end="", flush=True)
                else:
                    pygame.mixer.music.pause()
                    is_paused = True
                    print("\n[ Paused ]", end="", flush=True)
                continue

            # Calculate current absolute position in seconds
            ms_played_since_start = pygame.mixer.music.get_pos()
            if ms_played_since_start == -1: ms_played_since_start = 0
            
            # Absolute position = Offset + (Played_Time * Speed_Factor)
            current_abs_pos_sec = current_playback_start_offset + (ms_played_since_start / 1000.0) * playback_speed

            # --- Seek (FF/RW) using Arrows ---
            if command in ['right', 'left']:
                seek_amount = 5.0 if command == 'right' else -5.0
                new_pos = current_abs_pos_sec + seek_amount
                print(f"\n[ Seeking to {int(new_pos // 60)}:{int(new_pos % 60):02d} ]")
                
                if playback_speed == 1.0:
                    # Standard seek
                    play_at_position(new_pos)
                else:
                    # If speed is modified, seek in stretched domain
                    stretched_pos = new_pos / playback_speed
                    play_at_position(stretched_pos)
            
            # --- Speed Control ---
            elif command in ['speed_up', 'speed_down', 'speed_reset']:
                old_speed = playback_speed
                if command == 'speed_up': playback_speed = round(playback_speed + 0.1, 1)
                elif command == 'speed_down': playback_speed = max(0.5, round(playback_speed - 0.1, 1))
                elif command == 'speed_reset': playback_speed = 1.0

                if playback_speed != old_speed:
                    print(f"\n[ Changing speed to {playback_speed:.1f}x... Processing... ]")
                    pygame.mixer.music.stop()
                    
                    if playback_speed == 1.0:
                        # Revert to direct file reading (Fast)
                        pygame.mixer.music.load(filepath)
                        play_at_position(current_abs_pos_sec)
                    else:
                        # Process audio (Slow)
                        original = get_full_audio()
                        processed = apply_speed_change(original, playback_speed)
                        
                        # Export to temp file for pygame.mixer.music to stream
                        temp_wav = "temp_speed_audio.wav"
                        processed.export(temp_wav, format="wav")
                        
                        pygame.mixer.music.load(temp_wav)
                        stretched_start = current_abs_pos_sec / playback_speed
                        play_at_position(stretched_start)
                    
                    print("Done.")

        except queue.Empty:
            pass

        # --- UI Update (Text Sync) ---
        if pygame.mixer.music.get_busy() and not is_paused:
            ms_played = pygame.mixer.music.get_pos()
            if ms_played != -1:
                current_stretched_pos = current_playback_start_offset + (ms_played / 1000.0)
                real_pos_ms = current_stretched_pos * playback_speed * 1000.0
                
                # Find which sentence we are in
                found_index = -1
                for i, s in enumerate(sentences_data):
                    if s['start'] <= real_pos_ms < s['end']:
                        found_index = i
                        break
                
                # Only refresh screen if sentence changed
                if found_index != -1 and found_index != current_sentence_index:
                    current_sentence_index = found_index
                    os.system('cls' if os.name == 'nt' else 'clear')
                    print("\n" + "="*50)
                    print(f"--- Playing (Repeat=1) | Speed: {playback_speed:.1f}x ---")
                    print(f"Controls: (→) +5s, (←) -5s, (Space) Pause, (d/a/s) Speed")
                    print("-" * 50)
                    print(f"Sentence {current_sentence_index + 1}:")
                    print(f"{sentences_data[current_sentence_index]['text']}")
                    print("="*50)
        
        time.sleep(0.05)

    # Cleanup temp file if exists
    if os.path.exists("temp_speed_audio.wav"):
        try: os.remove("temp_speed_audio.wav")
        except: pass

def sentence_listening_practice(filepath, repeat_times=3, whisper_model="base"):
    if not os.path.exists(filepath):
        print(f"Error: Audio file not found at '{filepath}'")
        return

    try:
        # Step 1: Text Analysis (Whisper)
        # We always need this for text display, even if we play MP3 directly
        cache_filepath = f"{filepath}.{whisper_model}.whisper.cache"
        sentences_data = None
        
        # Check Cache
        if os.path.exists(cache_filepath):
            print(f"Cache detected. Loading text data from '{cache_filepath}'...")
            with open(cache_filepath, 'rb') as f:
                loaded_data = pickle.load(f)
            
            # Validate cache structure
            is_valid = False
            if loaded_data and isinstance(loaded_data, list):
                # Check if it has 'start' and 'end' keys which are needed for sync
                if 'start' in loaded_data[0] and 'end' in loaded_data[0]:
                    is_valid = True
            
            if is_valid:
                sentences_data = loaded_data
                print("Cache loaded successfully.")
            else:
                print("Cache outdated (missing timestamps). Re-analyzing...")

        # Run Whisper if no valid cache
        if not sentences_data:
            print("Analyzing audio with Whisper (GPU enabled if available)...")
            
            # GPU Check
            device = "cuda" if torch.cuda.is_available() else "cpu"
            print(f"Running on: {device.upper()}")
            
            model = whisper.load_model(whisper_model, device=device)
            print("Transcribing...")
            result = model.transcribe(filepath, language="ja")
            
            print("Processing data...")
            audio = AudioSegment.from_file(filepath)
            sentences_data = []
            for segment in result["segments"]:
                start_ms = int(segment['start'] * 1000)
                end_ms = int(segment['end'] * 1000)
                text = segment['text'].strip()
                if not text: continue
                
                # We store audio segment ONLY for repeat_times > 1 mode
                # For repeat_times == 1, we don't strictly need it, but good to have for unified cache
                audio_segment = audio[start_ms:end_ms]
                
                sentences_data.append({
                    'audio': audio_segment, 
                    'text': text,
                    'start': start_ms,
                    'end': end_ms
                })
            
            with open(cache_filepath, 'wb') as f:
                pickle.dump(sentences_data, f)

        # Step 2: Playback
        pygame.init()
        pygame.mixer.init()

        command_queue = queue.Queue()
        input_thread = threading.Thread(target=input_collector, args=(command_queue,), daemon=True)
        input_thread.start()

        if repeat_times == 1:
            run_continuous_mode(filepath, sentences_data, command_queue)
        else:
            run_sentence_playback_mode(sentences_data, repeat_times, command_queue)

    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()
    finally:
        pygame.quit()

if __name__ == '__main__':
    audio_file = './23_12_n3/3_2.mp3'
    sentence_listening_practice(
        filepath=audio_file, 
        repeat_times=1,  # Set to 1 for Continuous Mode
        whisper_model="large"
    )