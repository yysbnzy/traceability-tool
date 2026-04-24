# -*- coding: utf-8 -*-
"""
创建精简版zip文件
"""

import zipfile
import os

source_dir = r'C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool'
output_zip = r'C:\Users\Administrator\.openclaw\workspace\projects\traceability-tool\traceability_tool_release.zip'

# 需要包含的文件
files_to_include = [
    'run_analysis_fixed_v3.py',
    'run_analysis_test.py',
    'traceability_v2_lite.py',
    'create_test_doc.py',
    'check_duplicates.py',
    'gui.py',
    'dist/run_analysis_fixed_v3.exe',  # 打包好的EXE
]

with zipfile.ZipFile(output_zip, 'w', zipfile.ZIP_DEFLATED) as zipf:
    for file in files_to_include:
        file_path = os.path.join(source_dir, file)
        if os.path.exists(file_path):
            arcname = os.path.basename(file_path)
            zipf.write(file_path, arcname)
            print(f'Added: {arcname}')
        else:
            print(f'Not found: {file}')

print(f'\nRelease zip created: {output_zip}')
print(f'Size: {os.path.getsize(output_zip) / 1024 / 1024:.2f} MB')
