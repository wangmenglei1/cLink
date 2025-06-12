#!/usr/bin/env python3
"""
机柜上架功能测试脚本
验证批量上架API的数据一致性和可靠性
"""

import requests
import json
import time

BASE_URL = "http://172.31.60.204:58000"

def get_data():
    """获取系统数据"""
    devices = requests.get(f"{BASE_URL}/get_device_instances").json()['device_instances']
    racks = requests.get(f"{BASE_URL}/get_racks").json()['racks']
    return devices, racks

def check_data_consistency():
    """运行数据一致性检查"""
    response = requests.post(f"{BASE_URL}/check_data_consistency")
    return response.json()

def batch_mount_test(rack_id, device_updates):
    """测试批量上架"""
    response = requests.post(f"{BASE_URL}/batch_rack_mount", json={
        'rack_id': rack_id,
        'device_updates': device_updates
    })
    return response.json()

def batch_unmount_test(device_ids, unmount_type='selected'):
    """测试批量下架"""
    response = requests.post(f"{BASE_URL}/batch_rack_unmount", json={
        'device_ids': device_ids,
        'unmount_type': unmount_type
    })
    return response.json()

def main():
    print("=== 机柜上架功能测试 ===\n")
    
    # 1. 获取初始数据
    print("1. 获取系统数据...")
    devices, racks = get_data()
    print(f"   - 设备总数: {len(devices)}")
    print(f"   - 机柜总数: {len(racks)}")
    
    # 找到未上架的设备和可用机柜
    unmounted_devices = [d for d in devices if not d.get('rack_id')]
    available_racks = [r for r in racks if len(r.get('devices', [])) < r['height_u']]
    
    print(f"   - 未上架设备: {len(unmounted_devices)}")
    print(f"   - 可用机柜: {len(available_racks)}")
    
    if not unmounted_devices or not available_racks:
        print("⚠️  没有足够的设备或机柜进行测试")
        return
    
    # 2. 数据一致性检查（前）
    print("\n2. 执行数据一致性检查（测试前）...")
    consistency_before = check_data_consistency()
    print(f"   - 不一致项: {consistency_before.get('inconsistencies_found', 0)}")
    print(f"   - 自动修复: {consistency_before.get('repairs_made', 0)}")
    
    # 3. 批量上架测试
    print("\n3. 执行批量上架测试...")
    test_rack = available_racks[0]
    test_devices = unmounted_devices[:3]  # 测试前3个设备
    
    device_updates = []
    for i, device in enumerate(test_devices):
        device_updates.append({
            'device_id': device['id'],
            'rack_u': 30 + i * 3  # 从U30开始，避免冲突
        })
    
    print(f"   - 目标机柜: {test_rack['name']}")
    print(f"   - 测试设备: {len(device_updates)}台")
    
    mount_result = batch_mount_test(test_rack['id'], device_updates)
    print(f"   - 上架结果: {mount_result}")
    
    # 4. 验证上架结果
    print("\n4. 验证上架结果...")
    devices_after_mount, racks_after_mount = get_data()
    
    success_count = 0
    for update in device_updates:
        device = next((d for d in devices_after_mount if d['id'] == update['device_id']), None)
        if device and device.get('rack_id') == test_rack['id'] and device.get('rack_u') == update['rack_u']:
            success_count += 1
            print(f"   ✓ {device['instance_name']} 成功上架到 U{device['rack_u']}")
        else:
            print(f"   ✗ {update['device_id']} 上架失败")
    
    print(f"   - 验证成功率: {success_count}/{len(device_updates)}")
    
    # 5. 数据一致性检查（后）
    print("\n5. 执行数据一致性检查（测试后）...")
    consistency_after = check_data_consistency()
    print(f"   - 不一致项: {consistency_after.get('inconsistencies_found', 0)}")
    print(f"   - 自动修复: {consistency_after.get('repairs_made', 0)}")
    
    # 6. 批量下架测试
    print("\n6. 执行批量下架测试...")
    mounted_device_ids = [update['device_id'] for update in device_updates[:2]]  # 下架前2个
    
    unmount_result = batch_unmount_test(mounted_device_ids)
    print(f"   - 下架结果: {unmount_result}")
    
    # 7. 验证下架结果
    print("\n7. 验证下架结果...")
    devices_after_unmount, _ = get_data()
    
    unmount_success = 0
    for device_id in mounted_device_ids:
        device = next((d for d in devices_after_unmount if d['id'] == device_id), None)
        if device and not device.get('rack_id'):
            unmount_success += 1
            print(f"   ✓ {device['instance_name']} 成功下架")
        else:
            print(f"   ✗ {device_id} 下架失败")
    
    print(f"   - 下架成功率: {unmount_success}/{len(mounted_device_ids)}")
    
    # 8. 最终一致性检查
    print("\n8. 执行最终数据一致性检查...")
    consistency_final = check_data_consistency()
    print(f"   - 不一致项: {consistency_final.get('inconsistencies_found', 0)}")
    print(f"   - 自动修复: {consistency_final.get('repairs_made', 0)}")
    
    # 总结
    print(f"\n=== 测试总结 ===")
    print(f"✓ 批量上架成功率: {success_count}/{len(device_updates)}")
    print(f"✓ 批量下架成功率: {unmount_success}/{len(mounted_device_ids)}")
    print(f"✓ 数据一致性: {consistency_final.get('inconsistencies_found', 0)} 个问题")
    
    if success_count == len(device_updates) and unmount_success == len(mounted_device_ids) and consistency_final.get('inconsistencies_found', 0) == 0:
        print("🎉 所有测试通过！系统运行正常。")
    else:
        print("⚠️  发现问题，请检查系统状态。")

if __name__ == "__main__":
    main() 