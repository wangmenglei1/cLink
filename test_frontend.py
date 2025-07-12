#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
机房机柜管理系统前端功能测试脚本
测试所有主要功能的API接口
"""

import requests
import json
import time
from typing import Dict, List, Optional

# 配置
BASE_URL = "http://172.31.60.204:58000"
HEADERS = {"Content-Type": "application/json"}

class RackManagementTester:
    def __init__(self):
        self.base_url = BASE_URL
        self.headers = HEADERS
        self.test_results = []
        
    def log_test(self, test_name: str, success: bool, message: str, details: Optional[Dict] = None):
        """记录测试结果"""
        result = {
            "test_name": test_name,
            "success": success,
            "message": message,
            "details": details or {},
            "timestamp": time.strftime("%Y-%m-%d %H:%M:%S")
        }
        self.test_results.append(result)
        status = "✅ PASS" if success else "❌ FAIL"
        print(f"{status} {test_name}: {message}")
        if details:
            print(f"   详情: {details}")
        print()
    
    def make_request(self, method: str, endpoint: str, data: Optional[Dict] = None) -> Dict:
        """发送HTTP请求"""
        url = f"{self.base_url}{endpoint}"
        try:
            if method.upper() == "GET":
                response = requests.get(url, headers=self.headers)
            elif method.upper() == "POST":
                response = requests.post(url, json=data, headers=self.headers)
            elif method.upper() == "PUT":
                response = requests.put(url, json=data, headers=self.headers)
            elif method.upper() == "DELETE":
                response = requests.delete(url, headers=self.headers)
            else:
                raise ValueError(f"不支持的HTTP方法: {method}")
            
            # 检查响应内容类型
            content_type = response.headers.get('Content-Type', '')
            response_data = {}
            
            if 'application/json' in content_type:
                response_data = response.json() if response.content else {}
            elif 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' in content_type:
                # Excel文件响应
                response_data = {"file_size": len(response.content), "content_type": "excel"}
            elif 'text/html' in content_type:
                # HTML页面响应
                response_data = {"content_type": "html", "size": len(response.content)}
            else:
                # 其他类型响应
                response_data = {"content_type": content_type, "size": len(response.content)}
            
            return {
                "status_code": response.status_code,
                "data": response_data,
                "success": response.status_code < 400
            }
        except Exception as e:
            return {
                "status_code": 0,
                "data": {"error": str(e)},
                "success": False
            }
    
    def test_get_data_endpoints(self):
        """测试数据获取接口"""
        print("=" * 50)
        print("测试数据获取接口")
        print("=" * 50)
        
        endpoints = [
            ("/get_device_instances", "获取设备实例"),
            ("/get_racks", "获取机柜信息"),
            ("/get_device_types", "获取设备类型"),
            ("/get_rooms", "获取机房信息"),
            ("/get_device_groups", "获取设备组"),
        ]
        
        for endpoint, description in endpoints:
            result = self.make_request("GET", endpoint)
            data_count = "N/A"
            if result["success"] and isinstance(result["data"], dict):
                # 尝试获取各种可能的数据字段
                for key in ["device_instances", "racks", "device_types", "rooms", "device_groups"]:
                    if key in result["data"] and isinstance(result["data"][key], list):
                        data_count = len(result["data"][key])
                        break
            elif result["success"] and isinstance(result["data"], list):
                data_count = len(result["data"])
            
            self.log_test(
                f"GET {endpoint}",
                result["success"],
                f"{description} - 状态码: {result['status_code']}",
                {"data_count": data_count}
            )
    
    def test_device_management(self):
        """测试设备管理功能"""
        print("=" * 50)
        print("测试设备管理功能")
        print("=" * 50)
        
        # 获取现有设备
        devices_result = self.make_request("GET", "/get_device_instances")
        if not devices_result["success"]:
            self.log_test("设备管理预检", False, "无法获取设备列表")
            return
        
        devices = devices_result["data"].get("device_instances", [])
        if not devices:
            self.log_test("设备管理预检", False, "没有可用的设备进行测试")
            return
        
        # 测试创建设备实例
        test_device_data = {
            "deviceType": "TestServer",
            "instanceName": "测试服务器_001"
        }
        
        create_result = self.make_request("POST", "/add_device_instance", test_device_data)
        self.log_test(
            "创建设备实例",
            create_result["success"],
            f"创建测试设备 - 状态码: {create_result['status_code']}",
            create_result["data"]
        )
        
        # 如果创建成功，记录设备ID用于后续测试
        created_device_id = None
        if create_result["success"] and "id" in create_result["data"]:
            created_device_id = create_result["data"]["id"]
        
        # 测试更新设备实例
        if created_device_id:
            update_data = {
                "instance_name": "测试服务器_001_更新",
                "notes": "测试设备实例 - 已更新"
            }
            update_result = self.make_request("PUT", f"/api/device_instances/{created_device_id}", update_data)
            self.log_test(
                "更新设备实例",
                update_result["success"],
                f"更新测试设备 - 状态码: {update_result['status_code']}",
                update_result["data"]
            )
        
        # 测试删除设备实例（如果创建成功）
        if created_device_id:
            delete_result = self.make_request("DELETE", f"/api/device_instances/{created_device_id}")
            self.log_test(
                "删除设备实例",
                delete_result["success"],
                f"删除测试设备 - 状态码: {delete_result['status_code']}",
                delete_result["data"]
            )
    
    def test_rack_operations(self):
        """测试机柜操作功能"""
        print("=" * 50)
        print("测试机柜操作功能")
        print("=" * 50)
        
        # 获取机柜和设备信息
        racks_result = self.make_request("GET", "/get_racks")
        devices_result = self.make_request("GET", "/get_device_instances")
        
        if not racks_result["success"] or not devices_result["success"]:
            self.log_test("机柜操作预检", False, "无法获取机柜或设备信息")
            return
        
        racks = racks_result["data"].get("racks", [])
        devices = devices_result["data"].get("device_instances", [])
        
        if not racks or not devices:
            self.log_test("机柜操作预检", False, "没有可用的机柜或设备进行测试")
            return
        
        # 找到未上架的设备
        unmounted_devices = [d for d in devices if not d.get("rack_id")]
        mounted_devices = [d for d in devices if d.get("rack_id")]
        
        if not unmounted_devices:
            self.log_test("机柜操作预检", False, "没有未上架的设备进行上架测试")
        else:
            # 测试单设备上架
            test_device = unmounted_devices[0]
            test_rack = racks[0]
            
            mount_data = {
                "device_id": test_device["id"],
                "rack_id": test_rack["id"],
                "start_u": 1
            }
            
            mount_result = self.make_request("POST", "/rack_mount", mount_data)
            self.log_test(
                "单设备上架",
                mount_result["success"],
                f"设备 {test_device['instance_name']} 上架到机柜 {test_rack['name']} - 状态码: {mount_result['status_code']}",
                mount_result["data"]
            )
            
            # 如果上架成功，测试下架
            if mount_result["success"]:
                unmount_data = {"device_id": test_device["id"]}
                unmount_result = self.make_request("POST", "/rack_unmount", unmount_data)
                self.log_test(
                    "单设备下架",
                    unmount_result["success"],
                    f"设备 {test_device['instance_name']} 下架 - 状态码: {unmount_result['status_code']}",
                    unmount_result["data"]
                )
        
        # 测试批量上架
        if len(unmounted_devices) >= 2:
            batch_devices = unmounted_devices[:2]
            test_rack = racks[0]
            
            mount_requests = []
            for i, device in enumerate(batch_devices):
                mount_requests.append({
                    "device_id": device["id"],
                    "rack_id": test_rack["id"],
                    "start_u": i + 1
                })
            
            batch_mount_data = {"mount_requests": mount_requests}
            batch_mount_result = self.make_request("POST", "/multi_rack_batch_mount", batch_mount_data)
            self.log_test(
                "批量上架",
                batch_mount_result["success"],
                f"批量上架 {len(batch_devices)} 台设备 - 状态码: {batch_mount_result['status_code']}",
                batch_mount_result["data"]
            )
        
        # 测试批量下架
        if mounted_devices:
            # 测试选择性下架
            test_devices = mounted_devices[:2] if len(mounted_devices) >= 2 else mounted_devices
            device_ids = [d["id"] for d in test_devices]
            
            batch_unmount_data = {
                "unmount_type": "selected",
                "device_ids": device_ids
            }
            
            batch_unmount_result = self.make_request("POST", "/batch_rack_unmount", batch_unmount_data)
            self.log_test(
                "批量选择下架",
                batch_unmount_result["success"],
                f"批量下架 {len(device_ids)} 台设备 - 状态码: {batch_unmount_result['status_code']}",
                batch_unmount_result["data"]
            )
            
            # 测试整机柜下架
            if racks:
                test_rack = racks[0]
                rack_unmount_data = {
                    "unmount_type": "rack_all",
                    "rack_id": test_rack["id"]
                }
                
                rack_unmount_result = self.make_request("POST", "/batch_rack_unmount", rack_unmount_data)
                self.log_test(
                    "整机柜下架",
                    rack_unmount_result["success"],
                    f"机柜 {test_rack['name']} 整机柜下架 - 状态码: {rack_unmount_result['status_code']}",
                    rack_unmount_result["data"]
                )
    
    def test_ip_management(self):
        """测试IP管理功能"""
        print("=" * 50)
        print("测试IP管理功能")
        print("=" * 50)
        
        # 测试获取IP池
        ip_pools_result = self.make_request("GET", "/api/ip_pools")
        pools_count = "N/A"
        if ip_pools_result["success"]:
            if isinstance(ip_pools_result["data"], list):
                pools_count = len(ip_pools_result["data"])
            elif isinstance(ip_pools_result["data"], dict) and "pools" in ip_pools_result["data"]:
                pools_count = len(ip_pools_result["data"]["pools"])
        
        self.log_test(
            "获取IP池",
            ip_pools_result["success"],
            f"获取IP池信息 - 状态码: {ip_pools_result['status_code']}",
            {"pools_count": pools_count}
        )
        
        # 测试获取已分配子网
        subnets_result = self.make_request("GET", "/api/assigned_subnets")
        subnets_count = "N/A"
        if subnets_result["success"]:
            if isinstance(subnets_result["data"], list):
                subnets_count = len(subnets_result["data"])
            elif isinstance(subnets_result["data"], dict) and "assigned_subnets" in subnets_result["data"]:
                subnets_count = len(subnets_result["data"]["assigned_subnets"])
        
        self.log_test(
            "获取已分配子网",
            subnets_result["success"],
            f"获取子网信息 - 状态码: {subnets_result['status_code']}",
            {"subnets_count": subnets_count}
        )
        
        # 测试IP地址管理
        ip_addresses_result = self.make_request("GET", "/api/ip_addresses")
        addresses_count = "N/A"
        if ip_addresses_result["success"]:
            if isinstance(ip_addresses_result["data"], list):
                addresses_count = len(ip_addresses_result["data"])
            elif isinstance(ip_addresses_result["data"], dict) and "addresses" in ip_addresses_result["data"]:
                addresses_count = len(ip_addresses_result["data"]["addresses"])
        
        self.log_test(
            "获取IP地址",
            ip_addresses_result["success"],
            f"获取IP地址信息 - 状态码: {ip_addresses_result['status_code']}",
            {"addresses_count": addresses_count}
        )
    
    def test_connection_management(self):
        """测试连接管理功能"""
        print("=" * 50)
        print("测试连接管理功能")
        print("=" * 50)
        
        # 测试获取连接信息
        connections_result = self.make_request("GET", "/get_connections")
        connections_count = "N/A"
        if connections_result["success"]:
            if isinstance(connections_result["data"], list):
                connections_count = len(connections_result["data"])
            elif isinstance(connections_result["data"], dict) and "connections" in connections_result["data"]:
                connections_count = len(connections_result["data"]["connections"])
        
        self.log_test(
            "获取连接信息",
            connections_result["success"],
            f"获取连接信息 - 状态码: {connections_result['status_code']}",
            {"connections_count": connections_count}
        )
        
        # 测试获取端口模板
        port_templates_result = self.make_request("GET", "/get_device_port_templates")
        templates_count = "N/A"
        if port_templates_result["success"]:
            if isinstance(port_templates_result["data"], list):
                templates_count = len(port_templates_result["data"])
            elif isinstance(port_templates_result["data"], dict) and "templates" in port_templates_result["data"]:
                templates_count = len(port_templates_result["data"]["templates"])
        
        self.log_test(
            "获取端口模板",
            port_templates_result["success"],
            f"获取端口模板 - 状态码: {port_templates_result['status_code']}",
            {"templates_count": templates_count}
        )
    
    def test_error_handling(self):
        """测试错误处理"""
        print("=" * 50)
        print("测试错误处理")
        print("=" * 50)
        
        # 测试不存在的设备ID
        invalid_device_result = self.make_request("GET", "/get_device_instance/nonexistent_id")
        self.log_test(
            "不存在的设备ID",
            not invalid_device_result["success"],  # 期望失败
            f"访问不存在的设备 - 状态码: {invalid_device_result['status_code']}",
            invalid_device_result["data"]
        )
        
        # 测试无效的上架请求
        invalid_mount_data = {
            "device_id": "nonexistent_device",
            "rack_id": "nonexistent_rack",
            "start_u": -1
        }
        invalid_mount_result = self.make_request("POST", "/rack_mount", invalid_mount_data)
        self.log_test(
            "无效上架请求",
            not invalid_mount_result["success"],  # 期望失败
            f"无效上架请求 - 状态码: {invalid_mount_result['status_code']}",
            invalid_mount_result["data"]
        )
        
        # 测试无效的批量操作
        invalid_batch_data = {
            "unmount_type": "invalid_type",
            "device_ids": ["nonexistent1", "nonexistent2"]
        }
        invalid_batch_result = self.make_request("POST", "/batch_rack_unmount", invalid_batch_data)
        self.log_test(
            "无效批量下架",
            not invalid_batch_result["success"],  # 期望失败
            f"无效批量下架请求 - 状态码: {invalid_batch_result['status_code']}",
            invalid_batch_result["data"]
        )
    
    def test_rack_panel_functionality(self):
        """测试机架面板功能"""
        print("=" * 50)
        print("测试机架面板功能")
        print("=" * 50)
        
        # 获取机柜信息用于测试
        racks_result = self.make_request("GET", "/get_racks")
        if not racks_result["success"]:
            self.log_test("机架面板预检", False, "无法获取机柜信息")
            return
        
        racks = racks_result["data"].get("racks", [])
        if not racks:
            self.log_test("机架面板预检", False, "没有可用的机柜进行测试")
            return
        
        # 测试所有机架面板页面访问
        all_panels_result = self.make_request("GET", "/all_rack_panels")
        is_html_success = (all_panels_result["success"] and 
                          all_panels_result["data"].get("content_type") == "html")
        self.log_test(
            "所有机架面板页面访问",
            is_html_success,
            f"访问所有机架面板页面 - 状态码: {all_panels_result['status_code']}",
            {
                "content_type": all_panels_result["data"].get("content_type"),
                "size": all_panels_result["data"].get("size", 0)
            }
        )
        
        # 测试单个机架面板页面访问（取第一个机柜）
        test_rack = racks[0]
        single_panel_result = self.make_request("GET", f"/rack_panel/{test_rack['id']}")
        is_single_html_success = (single_panel_result["success"] and 
                                 single_panel_result["data"].get("content_type") == "html")
        self.log_test(
            "单个机架面板页面访问",
            is_single_html_success,
            f"访问机柜 {test_rack['name']} 的面板页面 - 状态码: {single_panel_result['status_code']}",
            {
                "rack_id": test_rack['id'], 
                "rack_name": test_rack['name'],
                "content_type": single_panel_result["data"].get("content_type"),
                "size": single_panel_result["data"].get("size", 0)
            }
        )
        
        # 测试不存在的机架面板访问
        invalid_panel_result = self.make_request("GET", "/rack_panel/invalid_rack_id")
        self.log_test(
            "无效机架面板访问",
            invalid_panel_result["status_code"] == 404,
            f"访问不存在的机架面板 - 状态码: {invalid_panel_result['status_code']}",
            {"expected_status": 404}
        )
        
        # 测试机架面板Excel导出功能
        export_result = self.make_request("GET", "/export_rack_panels_excel")
        is_excel_success = (export_result["success"] and 
                           export_result["data"].get("content_type") == "excel")
        self.log_test(
            "机架面板Excel导出",
            is_excel_success,
            f"导出机架面板Excel - 状态码: {export_result['status_code']}",
            {
                "content_type": export_result["data"].get("content_type"),
                "file_size": export_result["data"].get("file_size", 0)
            }
        )
        
        # 测试多个机架面板页面访问（如果有多个机柜）
        if len(racks) > 1:
            tested_racks = 0
            max_test_racks = min(3, len(racks))  # 最多测试3个机柜
            
            for i in range(1, max_test_racks):
                rack = racks[i]
                panel_result = self.make_request("GET", f"/rack_panel/{rack['id']}")
                if (panel_result["success"] and 
                    panel_result["data"].get("content_type") == "html"):
                    tested_racks += 1
            
            self.log_test(
                "多机架面板访问测试",
                tested_racks == (max_test_racks - 1),
                f"成功访问 {tested_racks}/{max_test_racks - 1} 个机架面板",
                {"tested_racks": tested_racks, "total_racks": len(racks)}
            )

    def run_all_tests(self):
        """运行所有测试"""
        print("🚀 开始运行机房机柜管理系统测试")
        print(f"测试目标: {self.base_url}")
        print("=" * 80)
        
        start_time = time.time()
        
        # 运行各项测试
        self.test_get_data_endpoints()
        self.test_device_management()
        self.test_rack_operations()
        self.test_ip_management()
        self.test_connection_management()
        self.test_error_handling()
        self.test_rack_panel_functionality()
        
        end_time = time.time()
        
        # 统计测试结果
        total_tests = len(self.test_results)
        passed_tests = sum(1 for r in self.test_results if r["success"])
        failed_tests = total_tests - passed_tests
        
        print("=" * 80)
        print("📊 测试结果统计")
        print("=" * 80)
        print(f"总测试数: {total_tests}")
        print(f"通过: {passed_tests} ✅")
        print(f"失败: {failed_tests} ❌")
        print(f"成功率: {(passed_tests/total_tests*100):.1f}%")
        print(f"测试耗时: {(end_time-start_time):.2f}秒")
        
        if failed_tests > 0:
            print("\n❌ 失败的测试:")
            for result in self.test_results:
                if not result["success"]:
                    print(f"  - {result['test_name']}: {result['message']}")
        
        print("\n" + "=" * 80)
        print("测试完成!")
        
        return {
            "total": total_tests,
            "passed": passed_tests,
            "failed": failed_tests,
            "success_rate": passed_tests/total_tests*100 if total_tests > 0 else 0,
            "duration": end_time - start_time,
            "results": self.test_results
        }
    
    def save_report(self, filename: str = "test_report.json"):
        """保存测试报告"""
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(self.test_results, f, ensure_ascii=False, indent=2)
        print(f"📄 测试报告已保存到: {filename}")

def main():
    """主函数"""
    tester = RackManagementTester()
    
    try:
        # 运行所有测试
        results = tester.run_all_tests()
        
        # 保存测试报告
        tester.save_report("rack_management_test_report.json")
        
        return results
        
    except KeyboardInterrupt:
        print("\n⚠️  测试被用户中断")
        return None
    except Exception as e:
        print(f"\n❌ 测试过程中发生错误: {e}")
        return None

if __name__ == "__main__":
    main() 