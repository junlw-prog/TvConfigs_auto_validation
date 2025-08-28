Update Note:

2025/8/28
# 進到有 configs/ 的上一層資料夾
python3 tvconfigs_path_check.py \
  --root ~/tvconfigs_home/tv109/kipling/configs

# 只想檢查特定副檔名
python3 tvconfigs_path_check.py --root ./configs --exts ini,bin,dat,img,xml

# 用於 CI（有缺檔就 fail）
python3 tvconfigs_path_check.py --root ./configs --fail-warning

=======
共有13種類別可以使用python判斷，可以檢視90%以上的場景，其他device的特例需要另外增加

<img width="1767" height="944" alt="image" src="https://github.com/user-attachments/assets/83745e33-bf55-4f03-9017-39609adef47f" />
<img width="1756" height="538" alt="image" src="https://github.com/user-attachments/assets/295e65af-6d89-4c35-b440-5ff0cb4d1764" />
