import os
import threading
import queue
import pickle
import io
from pydub import AudioSegment
import time

# 引入 Whisper 实现更精准的句子分割
import whisper

# 引入 numpy 和 librosa 以实现高质量的变速播放
import numpy as np
import librosa

# 为了更强大的音频控制（暂停/继续），我们用 pygame 替代 simpleaudio
import pygame
# 为了监听键盘事件，需要引入 pynput 库
from pynput import keyboard

def play_audio(segment):
    """
    使用 pygame 播放一个 pydub 的 AudioSegment 对象。
    """
    buffer = io.BytesIO()
    segment.export(buffer, format="wav")
    buffer.seek(0)
    sound = pygame.mixer.Sound(buffer)
    sound.play()

def apply_speed_change(segment, speed=1.0):
    """
    使用 librosa 实现高质量的倍速播放，保持音高不变。
    这是一个更可靠的方法，不依赖外部程序。
    """
    if speed == 1.0:
        return segment

    # 将 pydub segment 转换为 numpy 数组
    dtype_map = {1: np.int8, 2: np.int16, 4: np.int32}
    if segment.sample_width not in dtype_map:
        print(f"警告: 不支持的音频位宽 ({segment.sample_width})，将使用低质量变速。")
        return segment._spawn(segment.raw_data, overrides={"frame_rate": int(segment.frame_rate * speed)})

    dtype = dtype_map[segment.sample_width]
    samples = np.frombuffer(segment.raw_data, dtype=dtype)
    
    # 将样本归一化为浮点数 (-1.0 to 1.0)，供 librosa 使用
    samples_float = samples.astype(np.float32) / np.iinfo(dtype).max
    
    if segment.channels == 2:
        # librosa 需要 (channels, samples) 的形状
        samples_float = samples_float.reshape((-1, 2)).T

    # 使用 librosa进行时间拉伸
    stretched_samples = librosa.effects.time_stretch(y=samples_float, rate=speed)

    if segment.channels == 2:
        # 将形状转换回来
        stretched_samples = stretched_samples.T.flatten()

    # 将处理后的浮点数样本转换回原始整数类型
    y = (stretched_samples * np.iinfo(dtype).max).astype(dtype)

    # 创建新的 AudioSegment
    return AudioSegment(
        y.tobytes(),
        frame_rate=segment.frame_rate,
        sample_width=segment.sample_width,
        channels=segment.channels
    )


def input_collector(q):
    """
    在一个独立的线程中通过 pynput 监听键盘事件，并将指令放入队列。
    """
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
    """
    使用 Whisper 进行精准分割的逐句精听练习主函数。
    """
    if not os.path.exists(filepath):
        print(f"错误：找不到音频文件 '{filepath}'")
        return

    try:
        # --- 步骤 1: 使用 Whisper 进行分割，并利用缓存 ---
        cache_filepath = f"{filepath}.{whisper_model}.whisper.cache"
        
        if os.path.exists(cache_filepath):
            print(f"检测到 Whisper 缓存文件，正在从 '{cache_filepath}' 加载...")
            with open(cache_filepath, 'rb') as f:
                sentences_data = pickle.load(f)
            print("从缓存加载成功！")
        else:
            print("未找到缓存，首次加载将使用 Whisper 模型进行分析。")
            print(f"正在加载 Whisper 模型 '{whisper_model}'...")
            model = whisper.load_model(whisper_model)
            print("模型加载完毕，正在进行语音转录，请耐心等待...")
            result = model.transcribe(filepath, language="ja")
            
            print("转录完成，正在根据时间戳分割音频...")
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
                print(f"分割完成！正在创建缓存 '{cache_filepath}' 以便下次快速启动。")
                with open(cache_filepath, 'wb') as f:
                    pickle.dump(sentences_data, f)

        if not sentences_data:
            print("无法从音频中识别出任何句子。")
            return
            
        print(f"音频已成功分割成 {len(sentences_data)} 个句子。")
        print("\n*** 重要提示: 如果你更换了音频文件或Whisper模型，请手动删除 .cache 文件。 ***")

        # --- 初始化 Pygame Mixer ---
        pygame.init()
        pygame.mixer.init()

        command_queue = queue.Queue()
        input_thread = threading.Thread(target=input_collector, args=(command_queue,), daemon=True)
        input_thread.start()

        current_sentence_index = 0
        is_paused = False
        command_to_process = None
        playback_speed = 1.0

        print("\n练习开始！操作: (→)下一句, (←)上一句, (空格)暂停/继续, (r)重复, (q)退出")
        print("         速度控制: (d)加快0.1, (a)减慢0.1, (s)重置为正常速度")

        while 0 <= current_sentence_index < len(sentences_data):
            if command_to_process:
                if command_to_process == 'n':
                    if current_sentence_index < len(sentences_data) - 1: current_sentence_index += 1
                elif command_to_process == 'p':
                    if current_sentence_index > 0: current_sentence_index -= 1
                elif command_to_process == 'q':
                    print("练习结束，祝你学习进步！")
                    break
                command_to_process = None

            sentence_info = sentences_data[current_sentence_index]
            sentence_audio = sentence_info['audio']
            sentence_text = sentence_info['text']
            
            print("\n" + "="*50)
            print(f"--- 第 {current_sentence_index + 1}/{len(sentences_data)} 句 ---")
            print(f"  文本: {sentence_text}")
            print("="*50)

            playback_interrupted = False
            for i in range(repeat_times):
                segment_with_speed = apply_speed_change(sentence_audio, playback_speed)
                print(f"播放次数: {i + 1}/{repeat_times} (速度: {playback_speed:.1f}x)")
                play_audio(segment_with_speed)
                
                while pygame.mixer.get_busy() or is_paused:
                    try:
                        command = command_queue.get_nowait()
                        if command == 'toggle_pause':
                            if is_paused: pygame.mixer.unpause(); is_paused = False; print("[ 已恢复 ]", end="", flush=True)
                            else: pygame.mixer.pause(); is_paused = True; print("\n[ 已暂停 ]", end="", flush=True)
                        elif command in ['n', 'p', 'q', 'r']:
                            pygame.mixer.stop(); is_paused = False; command_to_process = command; playback_interrupted = True; print(f"\n接收到指令，中断播放。"); break
                        elif command in ['speed_up', 'speed_down', 'speed_reset']:
                            pygame.mixer.stop(); is_paused = False
                            if command == 'speed_up': playback_speed = round(playback_speed + 0.1, 1)
                            elif command == 'speed_down': playback_speed = max(0.5, round(playback_speed - 0.1, 1))
                            elif command == 'speed_reset': playback_speed = 1.0
                            print(f"\n[ 速度已更改为: {playback_speed:.1f}x ]"); command_to_process = 'r'; playback_interrupted = True; break
                    except queue.Empty:
                        time.sleep(0.1)

                if playback_interrupted: break
                
                # 在两次重复之间短暂暂停，听感更自然
                if i < repeat_times - 1:
                    time.sleep(0.4)
            
            # 如果所有重复播放都未被打断，则自动处理
            if not playback_interrupted:
                # 如果不是最后一句，自动播放下一句
                if current_sentence_index < len(sentences_data) - 1:
                    print("\n[ 自动播放下一句... ]")
                    time.sleep(0.5)
                    command_to_process = 'n'
                # 如果是最后一句，则等待用户指令
                else:
                    print("\n*** 已到达最后一句。请按 (←) 上一句, (r) 重复, (q) 退出或 (a/s/d) 调整速度。 ***")
                    command = command_queue.get()
                    if command in ['speed_up', 'speed_down', 'speed_reset']:
                        if command == 'speed_up': playback_speed = round(playback_speed + 0.1, 1)
                        elif command == 'speed_down': playback_speed = max(0.5, round(playback_speed - 0.1, 1))
                        elif command == 'speed_reset': playback_speed = 1.0
                        print(f"\n[ 速度已设定为: {playback_speed:.1f}x ]")
                        command_to_process = 'r'
                    else:
                        command_to_process = command

    except Exception as e:
        print(f"处理时发生错误: {e}")
    finally:
        pygame.quit()

if __name__ == '__main__':
    audio_file = './21_7/5_1.mp3'
    sentence_listening_practice(
        filepath=audio_file, 
        repeat_times=5,
        whisper_model="base" # 可选: "tiny", "base", "small", "medium"
    )

