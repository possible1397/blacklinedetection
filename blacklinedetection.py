import os
import cv2
import numpy as np
import pandas as pd
import tkinter as tk
from tkinter import messagebox

# 設定要處理的資料夾路徑為當前工作目錄
folder_path = os.getcwd()

# Excel 文件名稱和圖像資料夾名稱生成函數
def generate_unique_filename(base_filename, extension=None):
    counter = 1
    if extension:
        new_filename = f"{base_filename}.{extension}"
    else:
        new_filename = base_filename

    while os.path.exists(new_filename):
        if extension:
            new_filename = f"{base_filename}_{counter}.{extension}"
        else:
            new_filename = f"{base_filename}_{counter}"
        counter += 1
    return new_filename

# 生成唯一的Excel文件名稱
excel_base_filename = os.path.join(folder_path, "black_line_detection_summary")
excel_filename = generate_unique_filename(excel_base_filename, "xlsx")

# 生成唯一的幀圖像儲存資料夾名稱
output_dir = generate_unique_filename("output_frames")
os.makedirs(output_dir)

# 黑色判定閾值 (設為1, 表示幾乎全黑)
threshold = 1

# 初始化Excel數據列表
excel_data = []

# 遍歷資料夾中的所有.mp4檔案
for video_file in os.listdir(folder_path):
    if video_file.endswith(".mp4"):
        video_path = os.path.join(folder_path, video_file)
        cap = cv2.VideoCapture(video_path)

        frame_count = 0
        black_line_frames = []
        first_black_line_frame = None

        # 創建影片專屬的資料夾
        video_output_dir = os.path.join(output_dir, os.path.splitext(video_file)[0])
        os.makedirs(video_output_dir, exist_ok=True)

        while cap.isOpened():
            ret, frame = cap.read()
            if not ret:
                break

            frame_count += 1
            black_line_detected = False
            min_val = None
            std_dev = None

            for i in range(540, 551):
                line = frame[i, :, :]
                min_val = np.min(line)
                std_dev = np.std(line)

                if min_val < threshold and std_dev < 20:
                    frame[i, :, :] = [0, 0, 255]  # BGR模式下的紅色
                    black_line_detected = True
                    # 跳出循环，防止重复绘制多个值
                    break

            if black_line_detected:
                # 在图像上只显示一次 min_val 和 std_dev 的值
                if min_val is not None and std_dev is not None:
                    # 格式化數字，只保留一位小數
                    min_val_text = f"min_val: {min_val:.1f}"
                    std_dev_text = f"std_dev: {std_dev:.1f}"

                    # 使用紅色字體繪製文本
                    cv2.putText(frame, min_val_text, (10, 30), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)
                    cv2.putText(frame, std_dev_text, (10, 60), cv2.FONT_HERSHEY_SIMPLEX, 0.6, (0, 0, 255), 2)

                if first_black_line_frame is None:
                    first_black_line_frame = frame_count
                black_line_frames.append(frame_count)
                frame_filename = os.path.join(video_output_dir, f"{video_file}_frame_{frame_count}.png")
                cv2.imwrite(frame_filename, frame)
                print(f"Black line detected and marked in frame {frame_count} of {video_file}, saved as {frame_filename}")

        cap.release()

        video_name = os.path.basename(video_path)
        total_black_line_frames = len(black_line_frames)

        excel_data.append({
            "Video Name": video_name,
            "Total Frames": frame_count,  # 記錄總幀數
            "First Black Line Frame": first_black_line_frame,
            "Total Black Line Frames": total_black_line_frames
        })

# 建立DataFrame
df = pd.DataFrame(excel_data)

# 保存到Excel文件並捕捉可能的錯誤
try:
    df.to_excel(excel_filename, index=False)
    print(f"Summary saved to {excel_filename}")
except Exception as e:
    print(f"Failed to save Excel file: {e}")

# 彈出完成提示窗口
root = tk.Tk()
root.withdraw()  # 隱藏主窗口
messagebox.showinfo("Black Line Detection", "Finish！")
