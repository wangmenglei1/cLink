#!/usr/bin/env python3
# -*- coding: utf-8 -*-

def manual_fix_indentation():
    """手动修复所有已知的缩进问题"""
    
    with open('app.py', 'r', encoding='utf-8') as f:
        content = f.read()
    
    # 定义需要修复的具体行和正确的缩进
    fixes = [
        # 第2094行：if语句后的return应该缩进
        ('            return jsonify({\'message\': \'缺少设备ID列表或机柜分配信息\'}), 400', 
         '                return jsonify({\'message\': \'缺少设备ID列表或机柜分配信息\'}), 400'),
        
        # 其他可能的缩进问题
        ('        device_ids = data.get(\'device_ids\', [])\n        rack_assignments = data.get(\'rack_assignments\', [])\n        \n        if not device_ids or not rack_assignments:\n            return jsonify',
         '            device_ids = data.get(\'device_ids\', [])\n            rack_assignments = data.get(\'rack_assignments\', [])\n        \n            if not device_ids or not rack_assignments:\n                return jsonify'),
    ]
    
    # 应用修复
    for old, new in fixes:
        if old in content:
            content = content.replace(old, new)
            print(f"✓ 修复了缩进问题")
        else:
            print(f"✗ 未找到需要修复的内容")
    
    # 写回文件
    with open('app.py', 'w', encoding='utf-8') as f:
        f.write(content)
    
    print("手动修复完成！")

if __name__ == "__main__":
    manual_fix_indentation() 