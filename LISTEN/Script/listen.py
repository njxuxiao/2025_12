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
                    'r': 'r', 'q': 'q',
                    'a': 'speed_down', 's': 'speed_reset', 'd': 'speed_up'
                }[key.char]
        except AttributeError:
            if key == keyboard.Key.right: command = 'n'
            elif key == keyboard.Key.left: command = 'p'
            elif key == keyboard.Key.space: command = 'toggle_pause'
        if command:
            q.put(command)

    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()

def sentence_listening_practice(filepath, repeat_times=3, whisper_model="base"):
    if not os.path.exists(filepath):
        print(f"Error: Audio file not found at '{filepath}'")
        return

    try:
        cache_filepath = f"{filepath}.{whisper_model}.whisper.cache"
        
        if os.path.exists(cache_filepath):
            print(f"Whisper cache file detected. Loading from '{cache_filepath}'...")
            with open(cache_filepath, 'rb') as f:
                sentences_data = pickle.load(f)
            print("Successfully loaded from cache!")
        else:
            print("Cache not found. Analyzing audio with Whisper on first run.")
            print(f"Loading Whisper model '{whisper_model}'...")
            model = whisper.load_model(whisper_model)
            print("Model loaded. Transcribing audio, please wait...")
            result = model.transcribe(filepath, language="ja")
            
            print("Transcription complete. Splitting audio by timestamps...")
            audio = AudioSegment.from_file(filepath)
            sentences_data = []
            for segment in result["segments"]:
                start_ms = int(segment['start'] * 1000)
                end_ms = int(segment['end'] * 1000)
                text = segment['text'].strip()
                if not text: continue
                
                audio_segment = audio[start_ms:end_ms]
                sentences_data.append({'audio': audio_segment, 'text': text})
            
            if sentences_data:
                print(f"Splitting complete! Creating cache file '{cache_filepath}' for faster startup next time.")
                with open(cache_filepath, 'wb') as f:
                    pickle.dump(sentences_data, f)

        if not sentences_data:
            print("Could not detect any sentences in the audio.")
            return
            
        print(f"Audio successfully split into {len(sentences_data)} sentences.")
        print("\n*** IMPORTANT: If you change the audio file or the Whisper model, please delete the .cache file manually. ***")

        pygame.init()
        pygame.mixer.init()

        command_queue = queue.Queue()
        input_thread = threading.Thread(target=input_collector, args=(command_queue,), daemon=True)
        input_thread.start()

        current_sentence_index = 0
        is_paused = False
        command_to_process = None
        playback_speed = 1.0

        print("\nPractice started! Controls: (→) Next, (←) Previous, (Space) Pause/Resume, (r) Repeat, (q) Quit")
        print("         Speed Controls: (d) Speed Up +0.1, (a) Slow Down -0.1, (s) Reset Speed")

        while 0 <= current_sentence_index < len(sentences_data):
            if command_to_process:
                if command_to_process == 'n':
                    if current_sentence_index < len(sentences_data) - 1: current_sentence_index += 1
                elif command_to_process == 'p':
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
                        elif command in ['n', 'p', 'q', 'r']:
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
                
                if i < repeat_times - 1:
                    time.sleep(0.4)
            
            if not playback_interrupted:
                if current_sentence_index < len(sentences_data) - 1:
                    print("\n[ Auto-playing next sentence... ]")
                    time.sleep(0.5)
                    command_to_process = 'n'
                else:
                    print("\n*** Reached the last sentence. Press (←) Previous, (r) Repeat, (q) Quit or (a/s/d) to adjust speed. ***")
                    command = command_queue.get()
                    if command in ['speed_up', 'speed_down', 'speed_reset']:
                        if command == 'speed_up': playback_speed = round(playback_speed + 0.1, 1)
                        elif command == 'speed_down': playback_speed = max(0.5, round(playback_speed - 0.1, 1))
                        elif command == 'speed_reset': playback_speed = 1.0
                        print(f"\n[ Speed set to: {playback_speed:.1f}x ]")
                        command_to_process = 'r'
                    else:
                        command_to_process = command

    except Exception as e:
        print(f"An error occurred during processing: {e}")
    finally:
        pygame.quit()

if __name__ == '__main__':
    audio_file = './21_7/5_1.mp3'
    sentence_listening_practice(
        filepath=audio_file, 
        repeat_times=5,
        whisper_model="base"
    )

