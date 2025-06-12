#!/usr/bin/env python3
"""
数据一致性修复脚本
彻底解决机柜设备列表与设备rack_id不同步的问题
"""

import json
import os

def load_data(filepath):
    """加载JSON数据"""
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"Error loading {filepath}: {e}")
        return []

def save_data(data, filepath):
    """保存JSON数据"""
    try:
        with open(filepath, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=4, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error saving {filepath}: {e}")
        return False

def fix_data_consistency():
    """修复数据一致性问题"""
    print("=== 开始数据一致性修复 ===\n")
    
    # 加载数据
    device_instances = load_data('data/device_instances.json')
    racks = load_data('data/racks.json')
    
    if not device_instances or not racks:
        print("❌ 数据加载失败！")
        return False
    
    print(f"加载成功：{len(device_instances)} 台设备，{len(racks)} 个机柜\n")
    
    # 显示修复前状态
    print("🔍 修复前的数据状态：")
    inconsistencies = []
    
    for rack in racks:
        rack_devices_list = rack.get('devices', [])
        actual_devices = [d for d in device_instances if d.get('rack_id') == rack['id']]
        actual_device_ids = [d['id'] for d in actual_devices]
        
        print(f"  机柜 {rack['name']}:")
        print(f"    - devices列表: {rack_devices_list}")
        print(f"    - 实际设备: {actual_device_ids}")
        
        if set(rack_devices_list) != set(actual_device_ids):
            inconsistencies.append(rack['name'])
            print(f"    ❌ 数据不一致！")
        else:
            print(f"    ✅ 数据一致")
    
    print(f"\n发现 {len(inconsistencies)} 个机柜数据不一致：{inconsistencies}\n")
    
    # 执行修复
    print("🔧 开始修复...")
    repairs_made = []
    
    # 修复1：清理机柜中不存在的设备ID
    for rack in racks:
        if 'devices' in rack:
            original_devices = rack['devices'][:]
            valid_devices = []
            
            for device_id in original_devices:
                device = next((d for d in device_instances if d['id'] == device_id), None)
                if device and device.get('rack_id') == rack['id']:
                    valid_devices.append(device_id)
                else:
                    repairs_made.append(f"从机柜 {rack['name']} 移除无效设备 {device_id}")
            
            rack['devices'] = valid_devices
    
    # 修复2：添加缺失的设备到机柜列表
    for device in device_instances:
        if device.get('rack_id'):
            rack = next((r for r in racks if r['id'] == device['rack_id']), None)
            if rack:
                if 'devices' not in rack:
                    rack['devices'] = []
                if device['id'] not in rack['devices']:
                    rack['devices'].append(device['id'])
                    repairs_made.append(f"添加设备 {device['instance_name']} 到机柜 {rack['name']} 列表")
            else:
                # 机柜不存在，清除设备的rack_id
                device['rack_id'] = None
                device['rack_u'] = None
                repairs_made.append(f"清除设备 {device['instance_name']} 的无效机柜分配")
    
    # 修复3：去重机柜设备列表
    for rack in racks:
        if 'devices' in rack:
            original_count = len(rack['devices'])
            rack['devices'] = list(set(rack['devices']))
            if len(rack['devices']) != original_count:
                repairs_made.append(f"去除机柜 {rack['name']} 的重复设备")
    
    print(f"共执行 {len(repairs_made)} 项修复：")
    for repair in repairs_made:
        print(f"  - {repair}")
    
    # 保存修复后的数据
    if save_data(device_instances, 'data/device_instances.json') and save_data(racks, 'data/racks.json'):
        print("\n✅ 数据保存成功！")
    else:
        print("\n❌ 数据保存失败！")
        return False
    
    # 显示修复后状态
    print("\n🎉 修复后的数据状态：")
    for rack in racks:
        rack_devices_list = rack.get('devices', [])
        actual_devices = [d for d in device_instances if d.get('rack_id') == rack['id']]
        actual_device_ids = [d['id'] for d in actual_devices]
        device_names = [d['instance_name'] for d in actual_devices]
        
        print(f"  机柜 {rack['name']}:")
        print(f"    - 设备列表: {rack_devices_list}")
        print(f"    - 实际设备: {actual_device_ids}")
        print(f"    - 设备名称: {device_names}")
        
        if set(rack_devices_list) == set(actual_device_ids):
            print(f"    ✅ 数据一致")
        else:
            print(f"    ❌ 仍然不一致")
    
    print(f"\n=== 数据一致性修复完成 ===")
    print(f"修复项目数：{len(repairs_made)}")
    return True

if __name__ == "__main__":
    fix_data_consistency() 