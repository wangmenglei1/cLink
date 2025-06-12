#!/usr/bin/env python3
"""
脚本：修正save_data函数调用的参数顺序
从 save_data(data, filepath) 改为 save_data(filepath, data)
"""

import re
import os

def fix_save_data_calls(file_path):
    """修正文件中的save_data调用"""
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 定义替换模式
        replacements = {
            'save_data(racks, RACKS_FILE)': 'save_data(RACKS_FILE, racks)',
            'save_data(device_instances, DEVICE_INSTANCES_FILE)': 'save_data(DEVICE_INSTANCES_FILE, device_instances)',
            'save_data(connections, CONNECTIONS_FILE)': 'save_data(CONNECTIONS_FILE, connections)',
            'save_data(device_types, DEVICE_TYPES_FILE)': 'save_data(DEVICE_TYPES_FILE, device_types)',
            'save_data(rooms, ROOMS_FILE)': 'save_data(ROOMS_FILE, rooms)',
        }
        
        # 应用替换
        modified = False
        for old_call, new_call in replacements.items():
            if old_call in content:
                content = content.replace(old_call, new_call)
                modified = True
                print(f"替换: {old_call} -> {new_call}")
        
        # 处理特殊情况：save_data(default_data, filepath)
        pattern = r'save_data\(default_data, filepath\)'
        if re.search(pattern, content):
            content = re.sub(pattern, 'save_data(filepath, default_data)', content)
            modified = True
            print("替换: save_data(default_data, filepath) -> save_data(filepath, default_data)")
        
        # 如果有修改，写回文件
        if modified:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"文件 {file_path} 已更新")
            return True
        else:
            print(f"文件 {file_path} 无需修改")
            return False
            
    except Exception as e:
        print(f"处理文件 {file_path} 时出错: {e}")
        return False

if __name__ == "__main__":
    # 修正app.py文件
    app_file = "app.py"
    if os.path.exists(app_file):
        success = fix_save_data_calls(app_file)
        if success:
            print("✅ save_data调用已成功修正")
        else:
            print("❌ 修正save_data调用时出现问题")
    else:
        print(f"❌ 文件 {app_file} 不存在") 