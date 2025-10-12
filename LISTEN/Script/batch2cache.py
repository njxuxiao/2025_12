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
    
    except Exception as e:
        print(f"An error occurred during processing: {e}")



if __name__ == '__main__':
    audio_path = './21_7/'

    for mp3 in os.listdir(audio_path):
        if (mp3.endswith('.mp3')):
            audio_file = audio_path + mp3

            if not os.path.exists(audio_file + ".large.whisper.cache"):
                sentence_listening_practice(
                    filepath=audio_file, 
                    repeat_times=3,
                    whisper_model="large"
                )
            print("---------{}.is ok------------".format(mp3))
