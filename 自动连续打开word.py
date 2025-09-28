import os
import tkinter as tk
from tkinter import filedialog
import subprocess
import glob
import platform
import time
import psutil  # 需要安装psutil库


def select_word_file():
    """弹出对话框让用户选择一个Word文件"""
    root = tk.Tk()
    root.withdraw()  # 隐藏主窗口

    # 设置文件选择对话框，只显示Word文件
    file_path = filedialog.askopenfilename(
        title="选择一个Word文件",
        filetypes=[("Word文件", "*.doc;*.docx"), ("所有文件", "*.*")]
    )

    return file_path


def get_all_word_files(directory):
    """获取目录下所有的Word文件并排序"""
    # 查找.doc和.docx文件
    doc_files = glob.glob(os.path.join(directory, "*.doc"))
    docx_files = glob.glob(os.path.join(directory, "*.docx"))

    # 合并并排序文件列表
    all_word_files = doc_files + docx_files
    # 按文件名排序，可根据需要修改排序方式
    all_word_files.sort(key=lambda x: os.path.basename(x))

    return all_word_files


def get_word_process_name():
    """根据操作系统返回Word进程名称"""
    if platform.system() == 'Windows':
        return "WINWORD.EXE"
    elif platform.system() == 'Darwin':  # macOS
        return "Microsoft Word"
    else:  # Linux
        return "soffice.bin"  # 假设使用LibreOffice


def open_file_and_wait(file_path):
    """打开文件并等待其关闭"""
    word_process_name = get_word_process_name()

    # 获取打开文件前的所有Word进程ID
    initial_word_pids = set()
    for proc in psutil.process_iter(['pid', 'name']):
        if proc.info['name'] and word_process_name.lower() in proc.info['name'].lower():
            initial_word_pids.add(proc.info['pid'])

    # 打开文件
    try:
        if platform.system() == 'Windows':
            subprocess.Popen(['start', '', file_path], shell=True)
        elif platform.system() == 'Darwin':  # macOS
            subprocess.Popen(['open', file_path])
        else:  # Linux
            subprocess.Popen(['xdg-open', file_path])
    except Exception as e:
        print(f"打开文件 {file_path} 时出错: {e}")
        return False

    # 等待新的Word进程启动
    time.sleep(2)

    # 找到新启动的Word进程
    new_pid = None
    for proc in psutil.process_iter(['pid', 'name']):
        if (proc.info['name'] and word_process_name.lower() in proc.info['name'].lower() and
                proc.info['pid'] not in initial_word_pids):
            new_pid = proc.info['pid']
            break

    if not new_pid:
        print(f"无法跟踪文件 {file_path} 的进程，将继续下一个文件")
        return True

    print(f"请查看文件: {os.path.basename(file_path)}")
    print("关闭文件后将自动打开下一个...")

    # 等待进程结束
    try:
        if new_pid:
            proc = psutil.Process(new_pid)
            proc.wait()  # 等待进程结束
        return True
    except psutil.NoSuchProcess:
        print("文件进程已结束")
        return True
    except KeyboardInterrupt:
        print("\n程序被用户终止")
        return False
    except Exception as e:
        print(f"监控文件进程时出错: {e}")
        return True


def main():
    print("Word文件顺序打开器 (从选定文件开始)")
    print("-" * 40)

    # 让用户选择一个Word文件
    selected_file = select_word_file()
    if not selected_file or not os.path.isfile(selected_file):
        print("未选择有效的文件，程序退出")
        return

    # 检查所选文件是否为Word文件
    ext = os.path.splitext(selected_file)[1].lower()
    if ext not in ['.doc', '.docx']:
        print("所选文件不是Word文件，程序退出")
        return

    # 获取文件所在目录
    directory = os.path.dirname(selected_file)

    # 获取该目录下所有Word文件并排序
    all_word_files = get_all_word_files(directory)

    if not all_word_files:
        print(f"在目录 {directory} 中未找到任何Word文件")
        return

    # 找到所选文件在列表中的位置（使用绝对路径确保匹配）
    selected_path = os.path.abspath(selected_file)
    selected_index = next(i for i, path in enumerate(all_word_files)
                          if os.path.abspath(path) == selected_path)

    # 计算文件位置（从1开始计数）
    file_position = selected_index + 1
    total_files = len(all_word_files)

    print(f"所选文件是该目录中的第 {file_position} 个Word文件（共 {total_files} 个）")
    print("\n将从该文件开始依次打开后续文件：")
    for i in range(selected_index, total_files):
        print(f"{i + 1}. {os.path.basename(all_word_files[i])}")

    print("\n开始打开文件，按Ctrl+C可随时终止程序")
    print("-" * 40)

    # 从所选文件开始依次打开后续文件
    for i in range(selected_index, total_files):
        file_path = all_word_files[i]
        current_position = i + 1

        print(f"\n正在打开第 {current_position}/{total_files} 个文件: {os.path.basename(file_path)}")

        # 打开文件并等待关闭
        if not open_file_and_wait(file_path):
            print("程序终止")
            return

    print("\n所有后续Word文件已处理完毕")


if __name__ == "__main__":
    main()
