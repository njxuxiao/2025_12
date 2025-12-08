import os
import subprocess
import sys

def find_and_execute_scripts():
    """
    查找并执行当前目录及其子目录下的所有 'txt.py' 文件。
    为每个子脚本设置正确的工作目录。
    """
    start_dir = os.getcwd()
    print(f"[*] Starting search in: {start_dir}\n")

    found_files = 0

    for dirpath, _, filenames in os.walk(start_dir):
        # 排除脚本自身的目录，避免无限循环（如果脚本也叫txt.py）
        if os.path.samefile(dirpath, start_dir) and "run_all.py" in filenames:
            # 你也可以在这里添加其他需要忽略的文件夹
            pass
            
        for filename in filenames:
            if filename == 'txt.py':
                found_files += 1
                
                # 构建文件的完整路径
                file_path = os.path.join(dirpath, filename)
                
                # 【核心修改】获取该脚本所在的目录
                script_directory = os.path.dirname(file_path)
                
                print(f"====== Executing: {file_path} ======")
                print(f"  -> Setting working directory to: {script_directory}") # 打印信息，方便调试

                try:
                    # 【核心修改】使用 cwd 参数来指定子进程的工作目录
                    result = subprocess.run(
                        [sys.executable, file_path],
                        capture_output=True,
                        text=True,
                        check=True,
                        encoding='utf-8',
                        cwd=script_directory  # <-- 关键的修改在这里！
                    )

                    if result.stdout:
                        print("--- Script Output ---")
                        print(result.stdout.strip())
                        print("---------------------")

                except subprocess.CalledProcessError as e:
                    print(f"[!] An error occurred while executing {file_path}.")
                    if e.stdout:
                        print("--- Output (stdout) ---")
                        print(e.stdout.strip())
                    if e.stderr:
                        print("--- Error Message (stderr) ---")
                        print(e.stderr.strip())
                        print("----------------------------")
                except Exception as e:
                    print(f"[!] An unexpected error occurred while executing {file_path}: {e}")

                print(f"====== Finished executing: {file_path} ======\n")

    if found_files == 0:
        print("[*] No files named 'txt.py' were found in any subdirectories.")
    else:
        print(f"[*] Search complete. Found and attempted to execute {found_files} script(s).")


if __name__ == "__main__":
    find_and_execute_scripts()