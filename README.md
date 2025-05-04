# 🎬 Black Line Detection from MP4 Frames

本專案用於自動分析資料夾中的 `.mp4` 影片，偵測畫面第 540~550 行是否出現近乎純黑的線條，並將有問題的幀截圖儲存下來，同時產出一份 Excel 報告。

## ✅ 功能特色

- 批次掃描所有 `.mp4` 檔
- 偵測畫面中是否有黑線（min_val < 1 且 std_dev < 20）
- 自動截圖並存放於各自影片資料夾
- 自動產生 `.xlsx` 報表記錄下列資訊：
  - 影片名稱
  - 總幀數
  - 首次出現黑線的幀數
  - 出現黑線的總幀數
- GUI 視窗提示完成訊息（使用 tkinter）

## 🛠 使用方式

1. 將本程式碼放入含有 `.mp4` 檔的資料夾中
2. 使用 Python 執行
3. 結果將會輸出在：
   - `output_frames/` 目錄
   - Excel 報表檔（自動命名）

## 📦 相依套件

- `opencv-python`
- `numpy`
- `pandas`
- `tkinter`（內建於 Windows）

可使用以下指令安裝必要套件：

```bash
pip install opencv-python numpy pandas
