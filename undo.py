import shutil
from pathlib import Path

# 请将整理文件夹里的 undo_log.txt 路径复制过来
log_path = input('请输入撤销日志文件路径 (undo_log.txt): ').strip().strip('"')
log_path = Path(log_path)
if not log_path.exists():
    print('找不到日志文件。')
    input('按回车键退出...')
    exit()

archive_root = log_path.parent

with open(log_path, 'r', encoding='utf-8') as f:
    lines = f.readlines()

for line in lines:
    src, dst = line.strip().split('|')
    src, dst = Path(src), Path(dst)
    if dst.exists():
        shutil.move(str(dst), str(src))
        print(f'已恢复：{dst.name} → {src.parent}')

# 删除空文件夹
for folder in archive_root.iterdir():
    if folder.is_dir() and folder.name != 'undo_log.txt' and not any(folder.iterdir()):
        folder.rmdir()
        print(f'已删除空文件夹：{folder.name}')

log_path.unlink()
print('撤回完成，日志已删除。')
input('按回车键退出...')