Update Note:

2025/8/28
# 進到有 configs/ 的上一層資料夾
python3 tvconfigs_path_check.py \
  --root ~/tvconfigs_home/tv109/kipling/configs

# 只想檢查特定副檔名
python3 tvconfigs_path_check.py --root ./configs --exts ini,bin,dat,img,xml

# 用於 CI（有缺檔就 fail）
python3 tvconfigs_path_check.py --root ./configs --fail-warning

