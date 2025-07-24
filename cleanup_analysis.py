#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
设备数据清理工具
分析和清理残留的设备实例数据
"""

import json
import os
from datetime import datetime

def analyze_device_data():
    """分析设备数据"""
    # 读取设备实例数据
    with open('data/device_instances.json', 'r', encoding='utf-8') as f:
        data = json.load(f)

    device_instances = data.get('device_instances', [])

    print('=== 设备实例数据分析 ===')
    print(f'总设备数: {len(device_instances)}')

    # 分类设备
    mounted_devices = []  # 已上架设备（有名称和机柜位置）
    named_only_devices = []  # 仅有名称的设备（残留数据）
    unnamed_devices = []  # 未命名设备（残留数据）

    for device in device_instances:
        has_name = device.get('instance_name') and device.get('instance_name').strip()
        has_rack = device.get('rack_id')
        
        if has_name and has_rack:
            mounted_devices.append(device)
        elif has_name and not has_rack:
            named_only_devices.append(device)
        else:
            unnamed_devices.append(device)

    print(f'已上架设备: {len(mounted_devices)} 台')
    print(f'仅有名称未上架设备（残留）: {len(named_only_devices)} 台')
    print(f'未命名设备（残留）: {len(unnamed_devices)} 台')

    print('\n=== 残留设备详情 ===')
    if named_only_devices:
        print('仅有名称未上架的设备:')
        for i, device in enumerate(named_only_devices[:10]):  # 显示前10个
            name = device.get('instance_name', '未知')
            dtype = device.get('device_type', '未知型号')
            print(f'  {i+1}. {name} ({dtype})')
        if len(named_only_devices) > 10:
            print(f'  ... 还有 {len(named_only_devices) - 10} 台设备')

    if unnamed_devices:
        print('\n未命名设备:')
        for i, device in enumerate(unnamed_devices[:5]):  # 显示前5个
            device_id = device.get('id', '未知')
            dtype = device.get('device_type', '未知型号')
            print(f'  {i+1}. ID: {device_id} ({dtype})')
        if len(unnamed_devices) > 5:
            print(f'  ... 还有 {len(unnamed_devices) - 5} 台设备')
    
    return mounted_devices, named_only_devices, unnamed_devices

def cleanup_residual_data():
    """清理残留数据"""
    mounted_devices, named_only_devices, unnamed_devices = analyze_device_data()
    
    total_residual = len(named_only_devices) + len(unnamed_devices)
    if total_residual == 0:
        print('\n没有发现残留数据，无需清理。')
        return
    
    print(f'\n发现 {total_residual} 台残留设备')
    print('这些设备是规则更改前的历史数据，现在无法在界面中正常显示和管理')
    
    response = input('\n是否要清理这些残留数据？(y/N): ').strip().lower()
    if response != 'y':
        print('已取消清理操作')
        return
    
    # 备份原文件
    backup_file = f'data/device_instances_backup_{datetime.now().strftime("%Y%m%d_%H%M%S")}.json'
    with open('data/device_instances.json', 'r', encoding='utf-8') as f:
        original_data = f.read()
    with open(backup_file, 'w', encoding='utf-8') as f:
        f.write(original_data)
    print(f'已备份原文件到: {backup_file}')
    
    # 保存清理后的数据（只保留已上架设备）
    cleaned_data = {
        "version": "2.0",
        "timestamp": datetime.now().isoformat(),
        "device_instances": mounted_devices
    }
    
    with open('data/device_instances.json', 'w', encoding='utf-8') as f:
        json.dump(cleaned_data, f, ensure_ascii=False, indent=2)
    
    print(f'\n清理完成！')
    print(f'保留设备: {len(mounted_devices)} 台')
    print(f'清理设备: {total_residual} 台')
    print(f'备份文件: {backup_file}')
    
    # 更新库存（将清理的设备加回库存）
    update_inventory_after_cleanup(named_only_devices + unnamed_devices)

def update_inventory_after_cleanup(removed_devices):
    """清理后更新库存"""
    try:
        # 读取当前库存
        with open('data/inventory.json', 'r', encoding='utf-8') as f:
            inventory_data = json.load(f)
        
        inventory = inventory_data.get('inventory', [])
        
        # 统计被清理设备的型号
        device_type_counts = {}
        for device in removed_devices:
            device_type = device.get('device_type', '')
            if device_type:
                device_type_counts[device_type] = device_type_counts.get(device_type, 0) + 1
        
        # 将清理的设备加回库存
        for device_type, count in device_type_counts.items():
            inventory.append({
                'id': f'cleanup_{datetime.now().strftime("%Y%m%d_%H%M%S")}_{device_type}',
                'device_type': device_type,
                'quantity': count,
                'notes': f'数据清理时自动加回库存 - 清理了{count}台残留设备',
                'created_time': datetime.now().isoformat(),
                'created_by': '系统清理工具',
                'operation_type': 'cleanup'
            })
        
        # 保存更新后的库存
        inventory_data['inventory'] = inventory
        inventory_data['timestamp'] = datetime.now().isoformat()
        
        with open('data/inventory.json', 'w', encoding='utf-8') as f:
            json.dump(inventory_data, f, ensure_ascii=False, indent=2)
        
        print(f'\n已将清理的设备加回库存:')
        for device_type, count in device_type_counts.items():
            print(f'  {device_type}: {count} 台')
    
    except Exception as e:
        print(f'\n库存更新失败: {e}')
        print('请手动检查库存数据')

if __name__ == '__main__':
    print('设备数据清理工具')
    print('=' * 50)
    
    # 先分析数据
    analyze_device_data()
    
    print('\n' + '=' * 50)
    
    # 询问是否要清理
    cleanup_residual_data() 