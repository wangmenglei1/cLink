#!/usr/bin/env python3
"""
创建测试数据用于验证多机柜批量上架功能
"""
import requests
import json
import time

BASE_URL = "http://172.31.60.204:58000"

def create_test_device_type():
    """创建测试设备类型"""
    device_type_data = {
        "deviceModelName": "TestServer",
        "height_u": 1,
        "interfaces": [
            {
                "type": "Ethernet",
                "prefix": "GigabitEthernet",
                "naming": "0/0/",
                "start_num": 1,
                "end_num": 4
            }
        ]
    }
    
    response = requests.post(f"{BASE_URL}/add_device_type", json=device_type_data)
    if response.status_code == 201:
        print(f"✅ 设备类型创建成功: TestServer")
        return "TestServer"
    else:
        print(f"❌ 设备类型创建失败: {response.text}")
        return None

def create_test_room():
    """创建测试机房"""
    room_data = {
        "name": "数据中心A"
    }
    
    response = requests.post(f"{BASE_URL}/add_room", json=room_data)
    if response.status_code == 201:
        room_info = response.json()
        print(f"✅ 机房创建成功: {room_info['room']['name']} (ID: {room_info['room']['id']})")
        return room_info['room']['id']
    else:
        print(f"❌ 机房创建失败: {response.text}")
        return None

def create_test_racks(room_id, count=3):
    """创建测试机柜"""
    rack_ids = []
    for i in range(1, count + 1):
        rack_data = {
            "name": f"机柜A{i:02d}",
            "room_id": room_id,
            "height_u": 42
        }
        
        response = requests.post(f"{BASE_URL}/add_rack", json=rack_data)
        if response.status_code == 201:
            rack_info = response.json()
            print(f"✅ 机柜创建成功: {rack_info['rack']['name']} (ID: {rack_info['rack']['id']})")
            rack_ids.append(rack_info['rack']['id'])
        else:
            print(f"❌ 机柜创建失败: {response.text}")
    
    return rack_ids

def create_test_devices(device_type_name, count=6):
    """创建测试设备"""
    device_ids = []
    for i in range(1, count + 1):
        device_data = {
            "id": f"test_server_{i:02d}",
            "instanceName": f"TestServer{i:02d}",
            "deviceType": device_type_name,
            "position": {"x": 100 + i * 50, "y": 100}
        }
        
        response = requests.post(f"{BASE_URL}/add_device_instance", json=device_data)
        if response.status_code == 201:
            device_info = response.json()
            print(f"✅ 设备创建成功: {device_info['instance']['instance_name']} (ID: {device_info['instance']['id']})")
            device_ids.append(device_info['instance']['id'])
        else:
            print(f"❌ 设备创建失败: {response.text}")
    
    return device_ids

def test_multi_rack_batch_mount(device_ids, rack_ids):
    """测试多机柜批量上架功能"""
    if len(device_ids) < 3 or len(rack_ids) < 3:
        print("❌ 测试数据不足，至少需要3个设备和3个机柜")
        return False
    
    print("\n🚀 开始测试多机柜批量上架功能...")
    
    # 创建多机柜批量上架请求
    mount_requests = []
    for i, rack_id in enumerate(rack_ids):
        # 每个机柜安装2个设备（如果有足够的设备）
        start_idx = i * 2
        end_idx = min(start_idx + 2, len(device_ids))
        
        devices_for_rack = []
        for j in range(start_idx, end_idx):
            devices_for_rack.append({
                "device_id": device_ids[j],
                "rack_u": j - start_idx + 1  # U位从1开始
            })
        
        if devices_for_rack:
            mount_requests.append({
                "rack_id": rack_id,
                "device_updates": devices_for_rack
            })
    
    print(f"📋 准备批量上架 {len(mount_requests)} 个机柜，共 {sum(len(req['device_updates']) for req in mount_requests)} 台设备")
    
    # 使用新的多机柜批量上架API
    multi_rack_data = {
        "mount_requests": mount_requests
    }
    
    print(f"\n📦 执行多机柜批量上架...")
    response = requests.post(f"{BASE_URL}/multi_rack_batch_mount", json=multi_rack_data)
    
    if response.status_code in [200, 206]:  # 全部成功或部分成功
        result = response.json()
        print(f"✅ 多机柜上架完成: {result.get('message', '成功')}")
        print(f"📊 成功率: {result.get('success_rate', 'N/A')}")
        print(f"📊 机柜: {result.get('successful_racks', 0)}/{result.get('total_racks', 0)} 成功")
        print(f"📊 设备: {result.get('successful_devices', 0)}/{result.get('total_devices', 0)} 成功")
        
        if result.get('errors'):
            print("⚠️ 部分错误:")
            for error in result['errors']:
                print(f"   - {error}")
        
        return result.get('successful_racks', 0) == result.get('total_racks', 0)
    else:
        print(f"❌ 多机柜上架失败: {response.text}")
        return False

def main():
    print("🔧 开始创建测试数据...")
    
    # 创建设备类型
    device_type_name = create_test_device_type()
    if not device_type_name:
        return
    
    # 创建机房
    room_id = create_test_room()
    if not room_id:
        return
    
    # 创建机柜
    rack_ids = create_test_racks(room_id, 3)
    if len(rack_ids) < 3:
        print("❌ 机柜创建不足")
        return
    
    # 创建设备
    device_ids = create_test_devices(device_type_name, 6)
    if len(device_ids) < 6:
        print("❌ 设备创建不足")
        return
    
    # 测试多机柜批量上架
    success = test_multi_rack_batch_mount(device_ids, rack_ids)
    
    if success:
        print("\n🎉 多机柜批量上架功能测试通过！")
    else:
        print("\n⚠️ 多机柜批量上架功能测试失败！")

if __name__ == "__main__":
    main() 