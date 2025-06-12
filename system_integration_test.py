#!/usr/bin/env python3
"""
系统集成测试脚本 - 机柜管理系统
测试所有上架和下架功能的可靠性和数据一致性
"""

import requests
import json
import time
import random

BASE_URL = "http://172.31.60.204:58000"

class RackManagementTester:
    def __init__(self):
        self.session = requests.Session()
        self.initial_data = None
        self.test_results = {
            'passed': 0,
            'failed': 0,
            'details': []
        }
    
    def log_test(self, test_name, result, details=""):
        """记录测试结果"""
        status = "✓ PASS" if result else "✗ FAIL"
        message = f"{status} {test_name}"
        if details:
            message += f" - {details}"
        print(f"   {message}")
        
        self.test_results['details'].append({
            'test': test_name,
            'result': result,
            'details': details
        })
        
        if result:
            self.test_results['passed'] += 1
        else:
            self.test_results['failed'] += 1
    
    def api_call(self, method, endpoint, data=None):
        """通用API调用"""
        try:
            url = f"{BASE_URL}{endpoint}"
            if method == 'GET':
                response = self.session.get(url)
            elif method == 'POST':
                response = self.session.post(url, json=data)
            elif method == 'PUT':
                response = self.session.put(url, json=data)
            
            return response.status_code, response.json()
        except Exception as e:
            return 500, {'error': str(e)}
    
    def get_system_data(self):
        """获取系统数据"""
        devices_status, devices_data = self.api_call('GET', '/get_device_instances')
        racks_status, racks_data = self.api_call('GET', '/get_racks')
        
        if devices_status == 200 and racks_status == 200:
            return {
                'devices': devices_data.get('device_instances', []),
                'racks': racks_data.get('racks', [])
            }
        return None
    
    def check_data_consistency(self):
        """数据一致性检查"""
        status, data = self.api_call('POST', '/check_data_consistency')
        return status == 200, data
    
    def test_batch_mount(self, rack_id, device_updates):
        """测试批量上架"""
        status, data = self.api_call('POST', '/batch_rack_mount', {
            'rack_id': rack_id,
            'device_updates': device_updates
        })
        return status == 200, data
    
    def test_batch_unmount(self, device_ids=None, rack_id=None, unmount_type='selected'):
        """测试批量下架"""
        payload = {'unmount_type': unmount_type}
        if device_ids:
            payload['device_ids'] = device_ids
        if rack_id:
            payload['rack_id'] = rack_id
        
        status, data = self.api_call('POST', '/batch_rack_unmount', payload)
        return status == 200, data
    
    def run_comprehensive_test(self):
        """运行综合测试"""
        print("🚀 机柜管理系统 - 综合集成测试")
        print("=" * 50)
        
        # 1. 初始化测试
        print("\n📊 阶段 1: 系统初始化检查")
        initial_data = self.get_system_data()
        if not initial_data:
            print("❌ 无法获取系统数据，测试终止")
            return False
        
        devices = initial_data['devices']
        racks = initial_data['racks']
        
        self.log_test("系统数据获取", True, f"设备: {len(devices)}, 机柜: {len(racks)}")
        
        # 2. 数据一致性检查
        print("\n🔍 阶段 2: 数据一致性检查")
        consistency_ok, consistency_data = self.check_data_consistency()
        self.log_test("数据一致性检查", consistency_ok, 
                     f"不一致项: {consistency_data.get('inconsistencies_found', 0)}")
        
        # 3. 批量上架测试
        print("\n📦 阶段 3: 批量上架功能测试")
        unmounted_devices = [d for d in devices if not d.get('rack_id')]
        available_racks = [r for r in racks if len(r.get('devices', [])) < r['height_u']]
        
        if unmounted_devices and available_racks:
            # 测试小批量上架
            test_rack = available_racks[0]
            test_devices = unmounted_devices[:2]
            
            device_updates = []
            for i, device in enumerate(test_devices):
                device_updates.append({
                    'device_id': device['id'],
                    'rack_u': 20 + i * 3
                })
            
            mount_ok, mount_data = self.test_batch_mount(test_rack['id'], device_updates)
            self.log_test("小批量上架", mount_ok, 
                         f"设备数: {len(device_updates)}, 成功: {mount_data.get('success_count', 0)}")
            
            # 验证上架结果
            post_mount_data = self.get_system_data()
            if post_mount_data:
                mounted_count = 0
                for update in device_updates:
                    device = next((d for d in post_mount_data['devices'] if d['id'] == update['device_id']), None)
                    if device and device.get('rack_id') == test_rack['id']:
                        mounted_count += 1
                
                self.log_test("上架结果验证", mounted_count == len(device_updates),
                             f"预期: {len(device_updates)}, 实际: {mounted_count}")
        else:
            self.log_test("批量上架准备", False, "无可用设备或机柜")
        
        # 4. 批量下架测试
        print("\n📤 阶段 4: 批量下架功能测试")
        current_data = self.get_system_data()
        if current_data:
            mounted_devices = [d for d in current_data['devices'] if d.get('rack_id')]
            
            if mounted_devices:
                # 测试选择性下架
                test_devices = mounted_devices[:2]
                device_ids = [d['id'] for d in test_devices]
                
                unmount_ok, unmount_data = self.test_batch_unmount(device_ids=device_ids)
                self.log_test("选择性批量下架", unmount_ok,
                             f"设备数: {len(device_ids)}, 成功: {unmount_data.get('success_count', 0)}")
                
                # 验证下架结果
                post_unmount_data = self.get_system_data()
                if post_unmount_data:
                    unmounted_count = 0
                    for device_id in device_ids:
                        device = next((d for d in post_unmount_data['devices'] if d['id'] == device_id), None)
                        if device and not device.get('rack_id'):
                            unmounted_count += 1
                    
                    self.log_test("下架结果验证", unmounted_count == len(device_ids),
                                 f"预期: {len(device_ids)}, 实际: {unmounted_count}")
            else:
                self.log_test("批量下架准备", False, "无已上架设备")
        
        # 5. 整机柜操作测试
        print("\n🏗️ 阶段 5: 整机柜操作测试")
        current_data = self.get_system_data()
        if current_data:
            racks_with_devices = []
            for rack in current_data['racks']:
                devices_in_rack = [d for d in current_data['devices'] if d.get('rack_id') == rack['id']]
                if devices_in_rack:
                    racks_with_devices.append((rack, len(devices_in_rack)))
            
            if racks_with_devices:
                test_rack, device_count = racks_with_devices[0]
                
                # 整机柜下架
                rack_unmount_ok, rack_unmount_data = self.test_batch_unmount(
                    rack_id=test_rack['id'], unmount_type='rack_all'
                )
                self.log_test("整机柜下架", rack_unmount_ok,
                             f"机柜: {test_rack['name']}, 设备数: {device_count}")
            else:
                self.log_test("整机柜操作准备", False, "无包含设备的机柜")
        
        # 6. 压力测试
        print("\n⚡ 阶段 6: 系统压力测试")
        current_data = self.get_system_data()
        if current_data:
            unmounted_devices = [d for d in current_data['devices'] if not d.get('rack_id')]
            available_racks = [r for r in current_data['racks'] 
                             if len([d for d in current_data['devices'] if d.get('rack_id') == r['id']]) < r['height_u']]
            
            if len(unmounted_devices) >= 5 and available_racks:
                # 大批量上架测试
                test_rack = available_racks[0]
                test_devices = unmounted_devices[:5]
                
                device_updates = []
                for i, device in enumerate(test_devices):
                    device_updates.append({
                        'device_id': device['id'],
                        'rack_u': 1 + i * 2
                    })
                
                large_mount_ok, large_mount_data = self.test_batch_mount(test_rack['id'], device_updates)
                self.log_test("大批量上架", large_mount_ok,
                             f"设备数: {len(device_updates)}, 成功: {large_mount_data.get('success_count', 0)}")
            else:
                self.log_test("压力测试准备", False, "设备或机柜不足")
        
        # 7. 最终一致性检查
        print("\n🔧 阶段 7: 最终数据一致性检查")
        final_consistency_ok, final_consistency_data = self.check_data_consistency()
        self.log_test("最终一致性检查", final_consistency_ok,
                     f"不一致项: {final_consistency_data.get('inconsistencies_found', 0)}")
        
        # 测试总结
        print("\n" + "=" * 50)
        print("📋 测试总结")
        print(f"✅ 通过: {self.test_results['passed']}")
        print(f"❌ 失败: {self.test_results['failed']}")
        print(f"📊 成功率: {self.test_results['passed']/(self.test_results['passed']+self.test_results['failed'])*100:.1f}%")
        
        if self.test_results['failed'] == 0:
            print("🎉 所有测试通过！系统运行稳定。")
            return True
        else:
            print("⚠️ 发现问题，请检查失败的测试项。")
            return False

def main():
    tester = RackManagementTester()
    success = tester.run_comprehensive_test()
    return 0 if success else 1

if __name__ == "__main__":
    exit(main()) 