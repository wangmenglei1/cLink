#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import re

def fix_all_indentation():
    """一次性修复所有缩进问题"""
    
    with open('app.py', 'r', encoding='utf-8') as f:
        lines = f.readlines()
    
    # 记录修复的行数
    fixed_lines = []
    
    # 1. 修复第2094行 - if语句后的return缩进
    if 2093 < len(lines):  # 第2094行 (0-based index 2093)
        line = lines[2093]
        if line.strip().startswith('return') and not line.startswith('                '):
            lines[2093] = '                ' + line.strip() + '\n'
            fixed_lines.append(2094)
    
    # 2. 检查并修复所有常见的缩进问题
    for i in range(len(lines)):
        line = lines[i]
        prev_line = lines[i-1] if i > 0 else ""
        
        # 如果前一行以冒号结尾，当前行应该缩进
        if prev_line.strip().endswith(':'):
            if line.strip() and not line.startswith(' '):
                # 根据前一行的缩进级别决定当前行的缩进
                prev_indent = len(prev_line) - len(prev_line.lstrip())
                lines[i] = ' ' * (prev_indent + 4) + line.strip() + '\n'
                fixed_lines.append(i + 1)
        
        # 修复else语句后的缩进问题
        if prev_line.strip() == 'else:':
            if line.strip() and not line.startswith(' '):
                prev_indent = len(prev_line) - len(prev_line.lstrip())
                lines[i] = ' ' * (prev_indent + 4) + line.strip() + '\n'
                fixed_lines.append(i + 1)
    
    # 3. 特殊处理已知的问题行
    problem_areas = [
        # (起始行, 结束行, 目标缩进级别)
        (2090, 2100, 12),  # else块内容
    ]
    
    for start_line, end_line, target_indent in problem_areas:
        for i in range(start_line - 1, min(end_line, len(lines))):
            line = lines[i]
            if line.strip():  # 非空行
                current_indent = len(line) - len(line.lstrip())
                if current_indent != target_indent:
                    lines[i] = ' ' * target_indent + line.strip() + '\n'
                    fixed_lines.append(i + 1)
    
    # 写回文件
    with open('app.py', 'w', encoding='utf-8') as f:
        f.writelines(lines)
    
    print(f"缩进修复完成！")
    if fixed_lines:
        print(f"修复了 {len(fixed_lines)} 行: {fixed_lines[:10]}{'...' if len(fixed_lines) > 10 else ''}")
    else:
        print("没有发现需要修复的缩进问题")

if __name__ == "__main__":
    fix_all_indentation() 