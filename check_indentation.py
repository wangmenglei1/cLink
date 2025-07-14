#!/usr/bin/env python3
# -*- coding: utf-8 -*-

with open('app.py', 'r', encoding='utf-8') as f:
    content = f.read()
    
# 检查Tab和空格的使用情况
tab_count = content.count('\t')
lines = content.split('\n')
space_indents = 0
tab_indents = 0
mixed_indents = 0
problem_lines = []

for i, line in enumerate(lines):
    if line.strip():  # 非空行
        leading_spaces = len(line) - len(line.lstrip(' '))
        leading_tabs = len(line) - len(line.lstrip('\t'))
        
        if leading_tabs > 0 and leading_spaces > 0:
            mixed_indents += 1
            problem_lines.append(f"第{i+1}行: 混合缩进 (Tab + 空格)")
        elif leading_tabs > 0:
            tab_indents += 1
        elif leading_spaces > 0:
            space_indents += 1

print(f'总Tab字符数: {tab_count}')
print(f'使用空格缩进的行数: {space_indents}')
print(f'使用Tab缩进的行数: {tab_indents}')
print(f'混合缩进的行数: {mixed_indents}')

if problem_lines:
    print('\n问题行详情:')
    for line in problem_lines[:10]:  # 只显示前10个问题行
        print(line)
    if len(problem_lines) > 10:
        print(f"... 还有 {len(problem_lines) - 10} 行问题")

# 检查常见的缩进问题模式
print('\n分析可能的问题原因:')
if mixed_indents > 0:
    print("❌ 发现Tab和空格混用 - 这是最常见的缩进问题原因")
if tab_indents > 0 and space_indents > 0:
    print("❌ 文件中同时使用Tab和空格缩进")
if tab_count > 0:
    print("❌ 文件中包含Tab字符")
    
# 建议解决方案
print('\n建议的解决方案:')
print("1. 统一使用空格缩进 (Python推荐)")
print("2. 设置编辑器显示空白字符")
print("3. 配置编辑器自动转换Tab为空格")
print("4. 使用代码格式化工具 (如autopep8)") 