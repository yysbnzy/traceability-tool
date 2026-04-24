import os
import shutil
import glob

target_dir = r'C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool'
backup_dir = os.path.join(target_dir, 'backup')
dist_dir = os.path.join(target_dir, 'dist')

# 保留的最新版本文件
keep_files = {
    '.gitignore',
    'create_test_doc.py',
    'data_validator.py',
    'generate_docx.py',
    'GITHUB_UPLOAD_GUIDE.md',
    'md_to_docx.py',
    'README_v2.1.md',
    'requirements.txt',
    'traceability_v3_1.py',
}

# 遍历所有文件
for item in os.listdir(target_dir):
    item_path = os.path.join(target_dir, item)
    if os.path.isfile(item_path):
        if item in keep_files:
            continue
        # .spec 文件保留最新的
        if item.endswith('.spec') and 'v3.1' in item and '合并' in item:
            continue
        # .docx 使用说明保留
        if item.endswith('.docx') and '使用说明' in item:
            continue
        # .exe 移到 dist/
        if item.endswith('.exe'):
            shutil.move(item_path, os.path.join(dist_dir, item))
            print(f'Moved to dist/: {item}')
        else:
            # 其他移到 backup/
            shutil.move(item_path, os.path.join(backup_dir, item))
            print(f'Moved to backup/: {item}')

# 清空 build 目录
build_dir = os.path.join(target_dir, 'build')
if os.path.exists(build_dir):
    for item in os.listdir(build_dir):
        item_path = os.path.join(build_dir, item)
        if os.path.isfile(item_path):
            os.remove(item_path)
        elif os.path.isdir(item_path):
            shutil.rmtree(item_path)
    print('Cleaned build/')

print('Done!')
