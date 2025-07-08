# -*- coding: utf-8 -*-

from flask import Flask, render_template, request, jsonify, redirect, url_for, flash, send_file, g
import os
import threading
import json
import uuid
import threading
import shutil
import io
import time
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

app = Flask(__name__)

# Global variables
device_types = []
device_instances = []
connections = []
racks = []
rooms = []
device_groups = {}
device_port_templates = {}

# File path constants
DATA_DIR = 'data'
DEVICE_TYPES_FILE = os.path.join(DATA_DIR, 'device_types.json')
DEVICE_INSTANCES_FILE = os.path.join(DATA_DIR, 'device_instances.json')
CONNECTIONS_FILE = os.path.join(DATA_DIR, 'connections.json')
ROOMS_FILE = os.path.join(DATA_DIR, 'rooms.json')
RACKS_FILE = os.path.join(DATA_DIR, 'racks.json')
DEVICE_GROUPS_FILE = os.path.join(DATA_DIR, 'device_groups.json')
GROUP_TEMPLATES_FILE = os.path.join(DATA_DIR, 'group_templates.json')
PORT_TEMPLATES_FILE = os.path.join(DATA_DIR, 'port_templates.json')

# File lock objects
file_lock = threading.Lock()
file_locks = {
    'device_instances': threading.Lock(),
    'device_types': threading.Lock(),
    'connections': threading.Lock(),
    'racks': threading.Lock(),
    'rooms': threading.Lock(),
    'device_groups': threading.Lock(),
    'device_port_templates': threading.Lock()
}

# Interface configuration
INTERFACE_CONFIG = {
            "1G": {
        "interface_types": {
            "Optical": {"prefix": "Gi", "type": "optical"},
            "Electrical": {"prefix": "Gi", "type": "electrical"}
        },
        "default_interface": "Optical"
            },
            "10G": {
        "interface_types": {
            "Optical": {"prefix": "Te", "type": "optical"},
            "Electrical": {"prefix": "Te", "type": "electrical"}
        },
        "default_interface": "Optical"
            },
            "25G": {
        "interface_types": {
            "Optical": {"prefix": "TwentyFiveGigE", "type": "optical"},
            "Electrical": {"prefix": "TwentyFiveGigE", "type": "electrical"}
        },
        "default_interface": "Optical"
            },
            "40G": {
        "interface_types": {
            "Optical": {"prefix": "Fo", "type": "optical"},
            "Electrical": {"prefix": "Fo", "type": "electrical"}
        },
        "default_interface": "Optical"
            },
            "100G": {
        "interface_types": {
            "Optical": {"prefix": "HundredGigE", "type": "optical"},
            "Electrical": {"prefix": "HundredGigE", "type": "electrical"}
        },
        "default_interface": "Optical"
    },
    "Management": {
        "interface_types": {
            "Electrical": {"prefix": "M-GigabitEthernet", "type": "management"}
        },
        "default_interface": "Electrical"
    }
}

# Status mapping
status_map = {
    'on': '开启',
    'off': '关闭'
}

# --- Interface Configuration ---
# A structured configuration for defining interfaces based on rate and interface type.
# Note: INTERFACE_CONFIG is already defined above

# --- Data Storage --- #

# Function to ensure data directory exists
def ensure_data_dir():
    if not os.path.exists(DATA_DIR):
        os.makedirs(DATA_DIR)

# Function to load data from a JSON file
def load_data(filepath):
    """Load data from JSON file with corruption recovery"""
    ensure_data_dir()
    try:
        if os.path.exists(filepath):
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read().strip()
                if not content:
                    print(f"Warning: Empty file {filepath}. Using default empty structure.")
                    return get_default_data_structure(filepath)
                
                try:
                    data = json.loads(content)
                    return data
                except json.JSONDecodeError as e:
                    print(f"Warning: JSON decode error in {filepath}: {e}")
                    
                    # 尝试从备份恢复
                    backup_filepath = filepath + '.backup'
                    if os.path.exists(backup_filepath):
                        print(f"Attempting to restore from backup: {backup_filepath}")
                        try:
                            with open(backup_filepath, 'r', encoding='utf-8') as backup_f:
                                backup_data = json.load(backup_f)
                                # 恢复主文件
                                with open(filepath, 'w', encoding='utf-8') as main_f:
                                    json.dump(backup_data, main_f, ensure_ascii=False, indent=2)
                                print(f"Successfully restored {filepath} from backup")
                                return backup_data
                        except Exception as backup_error:
                            print(f"Backup restoration failed: {backup_error}")
                    
                    # 备份恢复失败，返回默认结构
                    print(f"Using default structure for corrupted file: {filepath}")
                    return get_default_data_structure(filepath)
        else:
            # 文件不存在，创建默认结构
            default_data = get_default_data_structure(filepath)
            save_data(filepath, default_data)
            return default_data
            
    except Exception as e:
        print(f'Error loading data from {filepath}: {e}')
        return get_default_data_structure(filepath)

def get_default_data_structure(filepath):
    """Get default data structure for different file types"""
    if 'device_instances' in filepath:
        return {'device_instances': []}
    elif 'racks' in filepath:
        return {'racks': []}
    elif 'rooms' in filepath:
        return {'rooms': []}
    elif 'device_types' in filepath:
        return {
            "version": "2.0",
            "timestamp": datetime.now().isoformat(),
            "device_types": [
                {
                    "name": "5130",
                    "height_u": 1,
                    "interfaces_definition": [
                        {
                            "type": "GigabitEthernet",
                            "prefix": "GE",
                            "naming": "1/0/",
                            "start_num": 1,
                            "end_num": 10
                        }
                    ]
                },
                {
                    "name": "HP",
                    "height_u": 2,
                    "interfaces_definition": [
                        {
                            "type": "GigabitEthernet",
                            "prefix": "GE",
                            "naming": "1/0/",
                            "start_num": 1,
                            "end_num": 1
                        }
                    ]
                }
            ]
        }
    elif 'connections' in filepath:
        return []  # 连接数据直接返回空列表
    elif 'device_groups' in filepath:
        return {'groups': []}
    else:
        return {}

# Function to save data to a JSON file
def save_data(filepath, data):
    """保存数据到文件，带有文件锁保护和版本信息"""
    try:
        # 确定锁的类型
        filename = os.path.basename(filepath)
        if 'device_instances' in filename:
            lock = file_locks['device_instances']
        elif 'racks' in filename:
            lock = file_locks['racks']
        elif 'rooms' in filename:
            lock = file_locks['rooms']
        elif 'device_types' in filename:
            lock = file_locks['device_types']
        elif 'connections' in filename:
            lock = file_locks['connections']
        elif 'device_groups' in filename:
            lock = file_locks['device_groups']
        else:
            lock = threading.Lock()
        
        with lock:
            # 使用临时文件确保原子性写
            temp_file = filepath + '.tmp'
            
            # 包装数据为新格式，添加版本信息和时间
            if 'device_instances' in filename:
                data_wrapper = {
                    "version": "2.0",
                    "timestamp": datetime.now().isoformat(),
                    "device_instances": data
                }
            elif 'racks' in filename:
                data_wrapper = {
                    "version": "2.0",
                    "timestamp": datetime.now().isoformat(),
                    "racks": data
                }
            elif 'rooms' in filename:
                data_wrapper = {
                    "version": "2.0",
                    "timestamp": datetime.now().isoformat(),
                    "rooms": data
                }
            elif 'device_types' in filename:
                data_wrapper = {
                    "version": "2.0",
                    "timestamp": datetime.now().isoformat(),
                    "device_types": data
                }
            elif 'connections' in filename:
                # 对于 connections.json，直接保存连接列表
                data_wrapper = data
            elif 'device_groups' in filename:
                data_wrapper = {
                    "version": "2.0",
                    "timestamp": datetime.now().isoformat(),
                    "groups": data
                }
            else:
                # 对于其他文件，直接保存原始数据
                data_wrapper = data
            
            # 写入临时文件
            with open(temp_file, 'w', encoding='utf-8') as f:
                json.dump(data_wrapper, f, ensure_ascii=False, indent=2)
            
            # 原子性替换
            if os.path.exists(temp_file):
                shutil.move(temp_file, filepath)
                
    except Exception as e:
        # 清理临时文件
        if os.path.exists(filepath + '.tmp'):
            os.remove(filepath + '.tmp')
        raise e

# Load data on application startup using before_request
@app.before_request
def load_all_data_before_request():
    global device_types, device_instances, connections, rooms, racks, device_groups, device_port_templates
    
    try:
        # Load all data files with corruption recovery
        device_types_data = load_data(DEVICE_TYPES_FILE)
        device_instances_data = load_data(DEVICE_INSTANCES_FILE)
        connections_data = load_data(CONNECTIONS_FILE)
        rooms_data = load_data(ROOMS_FILE)
        racks_data = load_data(RACKS_FILE)
        device_groups_data = load_data(DEVICE_GROUPS_FILE)
        
        # Extract data from new format
        device_types = device_types_data.get('device_types', [])
        device_instances = device_instances_data.get('device_instances', [])
        if isinstance(connections_data, dict) and 'connections' in connections_data:
            if isinstance(connections_data['connections'], dict) and 'connections' in connections_data['connections']:
                connections = connections_data['connections']['connections']
            else:
                connections = connections_data['connections']
        else:
            connections = connections_data if isinstance(connections_data, list) else []
        rooms = rooms_data.get('rooms', [])
        racks = racks_data.get('racks', [])
        device_groups = device_groups_data.get('groups', [])
        
        print(f"Data loaded - Rooms: {len(rooms)}, Racks: {len(racks)}, Devices: {len(device_instances)}, Types: {len(device_types)}, Connections: {len(connections)}, Groups: {len(device_groups)}")
        
    except Exception as e:
        print(f"Error loading data: {e}")
        device_types = []
        device_instances = []
        connections = []
        rooms = []
        racks = []
        device_groups = []

    # 加载端口模板数据
    with file_locks['device_port_templates']:
        try:
            with open(PORT_TEMPLATES_FILE, 'r', encoding='utf-8') as f:
                device_port_templates = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            device_port_templates = {
                "version": "1.0",
                "last_updated": datetime.now().strftime("%Y-%m-%d"),
                "templates": {}
            }
            save_device_port_templates()

def find_rack_by_id(rack_id):
    """Find rack by ID"""
    return next((rack for rack in racks if rack['id'] == rack_id), None)

def find_type_by_name(type_name):
    """Find device type by name"""
    return next((dt for dt in device_types if dt['name'] == type_name), None)

def get_occupied_slots(rack_id):
    """Get occupied U-slots in a rack, returns dict {U_number: device_name}"""
    occupied = {}
    rack = find_rack_by_id(rack_id)
    if not rack or 'devices' not in rack:
        return occupied

    for device_info in rack['devices']:
        device_id = device_info.get('instance_id')
        rack_u = device_info.get('rack_u')
        if not device_id or not rack_u:
            continue

        instance = find_instance_by_id(device_id)
        if not instance:
            continue

        device_type = find_type_by_name(instance.get('device_type'))
        height = device_type.get('height_u', 1) if device_type else 1
            
        for u in range(rack_u, rack_u + height):
            occupied[u] = instance.get('instance_name', 'Unknown Device')
    
    return occupied

# Function to find a device instance by its ID
def find_instance_by_id(instance_id):
    return next((inst for inst in device_instances if inst['id'] == instance_id), None)

# Function to find a connection by its details
def find_connection(source_id, target_id, source_port, target_port):
    for conn in connections:
        # Check both directions of the connection
        if (conn['source'] == source_id and conn['target'] == target_id and
            conn['source_port'] == source_port and conn['target_port'] == target_port) or \
           (conn['source'] == target_id and conn['target'] == source_id and
            conn['source_port'] == target_port and conn['target_port'] == source_port):
            return conn
    return None

# Function to update interface status of an instance
def update_interface_status(instance_id, port_name, status):
    instance = find_instance_by_id(instance_id)
    if instance:
        interface = next((iface for iface in instance['interfaces'] if iface['name'] == port_name), None)
        if interface:
            interface['status'] = status
            # print(f"Updated status of {port_name} on {instance_id} to {status}") # Debugging
            return True
    # print(f"Failed to update status of {port_name} on {instance_id}") # Debugging
    return False

# Function to perform the core delete device instance logic (used by route and cascade deletion)
def perform_delete_device_instance(instance_id):
    global device_instances, connections, racks

    # Find the instance to be deleted
    instance_to_delete = find_instance_by_id(instance_id)
    if not instance_to_delete:
        # Return False or raise an error if instance not found, depending on how this is called
        print(f"Warning: perform_delete_device_instance called for non-existent instance ID: {instance_id}")
        return False, [] # Return success status and list of updated instances

    # Collect connected port info from related connections BEFORE deleting connections
    ports_to_free = []
    connections_to_remove_details = []
    for conn in list(connections):
        if conn['source'] == instance_id:
            ports_to_free.append({'instance_id': conn['target'], 'port_name': conn['target_port']})
            connections_to_remove_details.append(conn) # Store connection details to free up ports correctly
        elif conn['target'] == instance_id:
            ports_to_free.append({'instance_id': conn['source'], 'port_name': conn['source_port']})
            connections_to_remove_details.append(conn) # Store connection details to free up ports correctly

    # Remove the device instance from the global list
    initial_instance_count = len(device_instances)
    device_instances = [inst for inst in device_instances if inst['id'] != instance_id]
    
    # Remove the instance from its associated rack's devices list if it was in a rack
    if 'rack_id' in instance_to_delete and instance_to_delete['rack_id'] is not None:
         rack = next((r for r in racks if r['id'] == instance_to_delete['rack_id']), None)
         if rack and 'devices' in rack:
             rack['devices'] = [dev_id for dev_id in rack['devices'] if dev_id != instance_id]
             save_data(RACKS_FILE, racks) # Save racks data after removing device from rack

    # Remove associated connections from the global list
    initial_connection_count = len(connections)
    connections = [conn for conn in connections if conn['source'] != instance_id and conn['target'] != instance_id]

    # Update interface status on connected instances
    updated_instances_data = []
    for port_info in ports_to_free:
         conn_detail = next((c for c in connections_to_remove_details if 
                             (c['source'] == instance_id and c['target'] == port_info['instance_id'] and c['target_port'] == port_info['port_name']) or
                             (c['target'] == instance_id and c['source'] == port_info['instance_id'] and c['source_port'] == port_info['port_name'])), None)
         
         if conn_detail: # Ensure we found the connection detail
              port_to_free_on_other_instance = port_info['port_name']
              
              if update_interface_status(port_info['instance_id'], port_to_free_on_other_instance, 'available'):
                  updated_instance = find_instance_by_id(port_info['instance_id'])
                  if updated_instance and updated_instance not in updated_instances_data:
                      updated_instances_data.append(updated_instance)

    # Save updated data to files
    save_data(DEVICE_INSTANCES_FILE, device_instances)
    if len(connections) != initial_connection_count:
         save_data(CONNECTIONS_FILE, connections)

    print(f"Performed delete for instance: {instance_id}. Removed {initial_instance_count - len(device_instances)} instance(s) and {initial_connection_count - len(connections)} connection(s).")

    # Return success status and list of updated instances data
    return True, updated_instances_data

# --- Routes --- #

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/rack_management')
def rack_management():
    return render_template('rack_management.html')

@app.route('/rack_panel/<string:rack_id>')
def single_rack_panel(rack_id):
    # Find the rack by ID
    rack = next((r for r in racks if r['id'] == rack_id), None)
    if not rack:
        return jsonify({'error': 'Not Found', 'message': f'Rack with ID {rack_id} not found'}), 404

    # Get device instances in this rack
    devices_in_rack = [inst for inst in device_instances if 'rack_id' in inst and inst['rack_id'] == rack_id]

    # Pass rack data, devices in rack, and all device types to the template
    return render_template('single_rack_panel.html', rack=rack, devices_in_rack=devices_in_rack, all_device_types=device_types)

@app.route('/all_rack_panels')
def all_rack_panels():
    from collections import defaultdict
    import re
    
    racks_by_row = defaultdict(list)
    for rack in racks:
        if rack['name']:
            row = rack['name'][0].upper()
            racks_by_row[row].append(rack)
    
    # 对每个排内的机柜按照编号排序
    def extract_number(rack_name):
        """从机柜名称中提取数字部分用于排序"""
        # 匹配机柜名称中的数字部分，如 A01 -> 1, A10 -> 10, North05 -> 5
        match = re.search(r'(\d+)', rack_name)
        return int(match.group(1)) if match else 0
    
    # 对每个排内的机柜按编号排�?
    for row in racks_by_row:
        racks_by_row[row].sort(key=lambda rack: extract_number(rack['name']))
    
    sorted_rows = sorted(racks_by_row.keys())
    return render_template('all_rack_panels.html',
        racks_by_row=racks_by_row,
        rows=sorted_rows,
        device_instances=device_instances,
        all_device_types=device_types,
        rooms=rooms
    )

@app.route('/get_device_types', methods=['GET'])
def get_device_types():
    """Returns a list of all defined device types."""
    try:
        load_all_data_before_request()
        
        # 确保 device_types 和 device_instances 已正确初始化
        global device_types, device_instances
        if not isinstance(device_types, list):
            print("Warning: device_types is not a list")
            device_types = []
        if not isinstance(device_instances, list):
            print("Warning: device_instances is not a list")
            device_instances = []
        
        # Create a set of used type names for efficient lookup
        used_types = {instance.get('device_type', '') for instance in device_instances if instance.get('device_type')}
        
        # Augment device types with usage info
        augmented_device_types = []
        for type_data in device_types:
            # Make a copy to avoid modifying the global list
            type_info = type_data.copy()
            type_info['is_in_use'] = type_info.get('name', '') in used_types
            augmented_device_types.append(type_info)

        return jsonify({'device_types': augmented_device_types}), 200
    except Exception as e:
        print(f"Error fetching device types: {str(e)}")
        # 返回一个空的设备类型列表而不是错误
        return jsonify({'device_types': []}), 200

@app.route('/api/interface_config', methods=['GET'])
def get_interface_config():
    """Provides the frontend with the structured interface configuration."""
    return jsonify(INTERFACE_CONFIG)

@app.route('/add_device_type', methods=['POST'])
def add_device_type():
    """Adds a new device type to the system using structured interface definitions."""
    try:
        load_all_data_before_request()
        data = request.get_json()
        if not data:
            return jsonify({'message': 'Invalid JSON data provided.'}), 400

        deviceModelName = data.get('deviceModelName', '').strip()
        height_u = data.get('height_u')
        # This will now contain structured groups with preset_key, naming_prefix, start, end.
        interfaces_definitions = data.get('interfaces', []) 
        
        if not deviceModelName or not isinstance(height_u, int) or height_u <= 0:
            return jsonify({'message': 'Device model name and a valid height (U) are required.'}), 400

        if any(d['name'].lower() == deviceModelName.lower() for d in device_types):
            return jsonify({'message': f"Device type '{deviceModelName}' already exists."}), 409

        # Generate the full interface list from structured definitions
        full_interfaces_list = []
        for interface_group in interfaces_definitions:
            rate = interface_group.get('rate')
            interface_type = interface_group.get('interface_type')
            naming_prefix = interface_group.get('naming_prefix', '1/0/')
            start = interface_group.get('start')
            end = interface_group.get('end')

            rate_config = INTERFACE_CONFIG.get(rate)
            if not rate_config:
                return jsonify({'message': f'无效的速率: {rate}'}), 400
            
            interface_config = rate_config['interface_types'].get(interface_type)
            if not interface_config:
                return jsonify({'message': f'速率 "{rate}" 不支持接口类型"{interface_type}"'}), 400

            if not all([isinstance(start, int), isinstance(end, int)]):
                return jsonify({'message': '接口组缺少参数或类型错误'}), 400
            if start > end:
                return jsonify({'message': f'接口"{rate} - {interface_type}" 的起始编号不能大于结束编号'}), 400

            base_prefix = interface_config['prefix']
            if_type = interface_config['type']

            for i in range(start, end + 1):
                interface_name = f"{base_prefix}{naming_prefix}{i}"
                full_interfaces_list.append({'name': interface_name, 'type': if_type, 'connected_to': None})

        new_device_type = {
            'name': deviceModelName,
            'height_u': height_u,
            'interfaces': full_interfaces_list,
            'definitions': interfaces_definitions # Store the structured definitions for editing
        }
        device_types.append(new_device_type)
        save_data(DEVICE_TYPES_FILE, device_types)
        
        print(f"[API] add_device_type: Added new device type '{deviceModelName}'")
        return jsonify({'message': 'Device type added successfully', 'device_type': new_device_type}), 201

    except Exception as e:
        print(f'Error adding device type: {e}')
        return jsonify({'message': 'Internal server error during device type addition'}), 500

@app.route('/update_device_type/<string:original_type_name>', methods=['PUT'])
def update_device_type(original_type_name):
    """Updates an existing device type using structured interface definitions."""
    try:
        load_all_data_before_request()

        # Find the target device type
        target_type = next((t for t in device_types if t['name'] == original_type_name), None)
        if not target_type:
            return jsonify({'message': '要更新的设备型号未找到'}), 404

        # Prevent editing if the type is in use
        if any(instance['device_type'] == original_type_name for instance in device_instances):
            return jsonify({'message': f'设备型号 "{original_type_name}" 正在使用中，无法修改'}), 400

        data = request.get_json()
        new_name = data.get('deviceModelName', '').strip()
        new_height = data.get('height_u')
        new_definitions = data.get('interfaces', [])

        # Validate new data
        if not new_name or not isinstance(new_height, int) or new_height <= 0:
            return jsonify({'message': '设备型号名称和有效高度为必填项'}), 400
        
        # Check for name conflict if the name is being changed
        if new_name != original_type_name and any(t['name'] == new_name for t in device_types):
            return jsonify({'message': f'设备型号名称 "{new_name}" 已存在'}), 409

        # Generate new full interface list from structured definitions
        new_full_interfaces_list = []
        for interface_group in new_definitions:
            rate = interface_group.get('rate')
            interface_type = interface_group.get('interface_type')
            naming_prefix = interface_group.get('naming_prefix', '1/0/')
            start = interface_group.get('start')
            end = interface_group.get('end')

            rate_config = INTERFACE_CONFIG.get(rate)
            if not rate_config:
                return jsonify({'message': f'无效的速率: {rate}'}), 400

            interface_config = rate_config['interface_types'].get(interface_type)
            if not interface_config:
                return jsonify({'message': f'速率 "{rate}" 不支持接口类型"{interface_type}"'}), 400

            if not all([isinstance(start, int), isinstance(end, int)]):
                return jsonify({'message': '接口组数据无效'}), 400
            
            base_prefix = interface_config['prefix']
            if_type = interface_config['type']

            for i in range(start, end + 1):
                new_full_interfaces_list.append({'name': f"{base_prefix}{naming_prefix}{i}", 'type': if_type, 'connected_to': None})

        # Update the device type
        target_type['name'] = new_name
        target_type['height_u'] = new_height
        target_type['interfaces'] = new_full_interfaces_list
        target_type['definitions'] = new_definitions

        save_data(DEVICE_TYPES_FILE, device_types)
        return jsonify({'message': '设备型号更新成功'}), 200

    except Exception as e:
        print(f"Error updating device type: {str(e)}")
        return jsonify({'message': '更新设备型号时发生内部错误'}), 500

@app.route('/delete_device_type/<string:device_type_name>', methods=['DELETE'])
def delete_device_type(device_type_name):
    """Deletes a device type from the system."""
    try:
        load_all_data_before_request()

        # Check if the device type is in use before deleting
        if any(instance['device_type'] == device_type_name for instance in device_instances):
            return jsonify({'message': f'设备型号 "{device_type_name}" 正在被一个或多个设备实例使用，无法删除'}), 400

        initial_count = len(device_types)
        # Using a global variable that is assumed to be loaded
        device_types[:] = [d for d in device_types if d['name'] != device_type_name]

        if len(device_types) < initial_count:
            save_data(DEVICE_TYPES_FILE, device_types)
            print(f'Deleted device type: {device_type_name}')
            return jsonify({'message': 'Device type deleted successfully'}), 200
        else:
            return jsonify({'message': f'Device type {device_type_name} not found'}), 404

    except Exception as e:
        print(f'Error deleting device type: {e}')
        return jsonify({'message': 'Internal server error during device type deletion'}), 500

@app.route('/get_device_instance/<string:instance_id>', methods=['GET'])
def get_device_instance(instance_id):
    instance = find_instance_by_id(instance_id)
    if instance:
        return jsonify({'instance': instance}), 200
    return jsonify({'error': 'Not Found', 'message': f'Device instance with ID {instance_id} not found'}), 404

@app.route('/add_device_instance', methods=['POST'])
def add_device_instance():
    """Adds a new device instance to the 'warehouse' without a position."""
    data = request.get_json()
    if not data:
        return jsonify({'message': 'Invalid JSON data received'}), 400

    instance_name = data.get('instanceName', '').strip()
    device_type_name = data.get('deviceType')
    
    # Core fields for a "warehouse" device. Position is no longer required.
    if not instance_name or not device_type_name:
         return jsonify({'message': 'instanceName and deviceType are required.'}), 400

    # Check for unique instance_name (case-insensitive)
    if any(inst['instance_name'].lower() == instance_name.lower() for inst in device_instances):
        return jsonify({'message': f'Device instance with name "{instance_name}" already exists.'}), 409

    device_type_definition = next((dt for dt in device_types if dt['name'] == device_type_name), None)

    if not device_type_definition:
        return jsonify({'message': f'Device type "{device_type_name}" not found.'}), 404

    # Create the interfaces for the new instance based on its type
    instance_interfaces = []
    # Correctly iterate over 'interfaces', use .get() for safety
    for iface_template in device_type_definition.get('interfaces', []):
        instance_interfaces.append({
            'name': iface_template['name'],
            'type': iface_template['type'],
            'status': 'available'
        })
    
    # Add a default management interface, ensuring it's always present
    has_mgmt_if = any(iface['name'] == 'M-GigabitEthernet0/0/0' for iface in instance_interfaces)
    if not has_mgmt_if:
        instance_interfaces.insert(0, {
            'name': 'M-GigabitEthernet0/0/0', 'type': 'management', 'status': 'available'
        })

    # Generate a new unique ID on the server
    instance_id = f'device_{uuid.uuid4().hex[:12]}'

    new_instance = {
        'id': instance_id,
        'instance_name': instance_name,
        'device_type': device_type_name,
        'interfaces': instance_interfaces,
        'position': None, # Explicitly null, as it's in the "warehouse"
        'rack_id': None,  # Not in a rack yet
        'rack_u': None,   # No U position yet
        'power_status': 'off'
    }
    device_instances.append(new_instance)
    save_data(DEVICE_INSTANCES_FILE, device_instances)

    return jsonify({'message': 'Device instance added successfully to warehouse', 'instance': new_instance}), 201

@app.route('/batch_add_device_instances', methods=['POST'])
def batch_add_device_instances():
    data = request.get_json()
    if not data:
        return jsonify({'message': '无效的JSON数据'}), 400

    device_type_name = data.get('deviceType')
    name_prefix = data.get('namePrefix', '').strip()
    start_number = data.get('start')
    end_number = data.get('end')

    # --- Validation ---
    if not all([device_type_name, name_prefix, isinstance(start_number, int), isinstance(end_number, int)]):
        return jsonify({'message': '设备型号、命名前缀、起始和结束编号均为必填项'}), 400
    
    if start_number > end_number:
        return jsonify({'message': '起始编号不能大于结束编号'}), 400

    if end_number - start_number + 1 > 200: # Limit batch size to prevent abuse/overload
        return jsonify({'message': '单次批量添加的设备数量不能超过200个'}), 400
    
    load_all_data_before_request()

    device_type_definition = next((dt for dt in device_types if dt['name'] == device_type_name), None)
    if not device_type_definition:
        return jsonify({'message': f'设备型号 "{device_type_name}" 未找到'}), 404

    existing_names = {inst['instance_name'].lower() for inst in device_instances}
    new_instances = []
    skipped_names = []

    for i in range(start_number, end_number + 1):
        instance_name = f"{name_prefix}{i:02d}"
        if instance_name.lower() in existing_names:
            skipped_names.append(instance_name)
            continue

        # Create interfaces for the new instance
        instance_interfaces = []
        for iface_template in device_type_definition.get('interfaces', []):
            instance_interfaces.append({
                'name': iface_template['name'],
                'type': iface_template['type'],
                'status': 'available'
            })
        
        new_instance = {
            'id': f'device_{uuid.uuid4().hex[:12]}',
            'instance_name': instance_name,
            'device_type': device_type_name,
            'interfaces': instance_interfaces,
            'position': None,
            'rack_id': None,
            'rack_u': None,
            'power_status': 'off'
        }
        new_instances.append(new_instance)

    if not new_instances:
        return jsonify({'message': f'未能添加任何新设备。以下名称均已存在 {", ".join(skipped_names)}'}), 409

    device_instances.extend(new_instances)
    save_data(DEVICE_INSTANCES_FILE, device_instances)

    message = f'成功添加 {len(new_instances)} 台设备'
    if skipped_names:
        message += f' 跳过 {len(skipped_names)} 台已存在的设备 {", ".join(skipped_names)}'
    
    return jsonify({'message': message, 'added_count': len(new_instances), 'skipped_count': len(skipped_names)}), 201

@app.route('/get_device_instances', methods=['GET'])
def get_device_instances():
    try:
        load_all_data_before_request()
        return jsonify({'device_instances': device_instances}), 200
    except Exception as e:
        print(f"Error getting device instances: {str(e)}")
        return jsonify({'message': '获取设备实例时发生错误'}), 500

@app.route('/update_device_instance_position/<string:instance_id>', methods=['PUT'])
def update_device_instance_position(instance_id):
    try:
        data = request.get_json()
        new_position = data.get('position')

        if not new_position or 'x' not in new_position or 'y' not in new_position:
            return jsonify({'message': 'Invalid position data provided'}), 400

        # Find the device instance by ID and update its position
        instance = find_instance_by_id(instance_id)
        if instance:
            instance['position'] = new_position
            save_data(DEVICE_INSTANCES_FILE, device_instances)
            print(f'Updated position for instance {instance_id}: {new_position}')
            return jsonify({'message': 'Position updated successfully', 'instance_id': instance_id, 'new_position': new_position}), 200
        else:
            return jsonify({'message': f'Device instance with ID {instance_id} not found'}), 404

    except Exception as e:
        print(f'Error updating device instance position: {e}')
        return jsonify({'message': 'Internal server error'}), 500

@app.route('/export_connections', methods=['GET'])
def export_connections():
    try:
        # Ensure the latest data is loaded (though before_request should handle this)
        load_all_data_before_request()

        # Prepare connection data for export
        exported_data = []
        for conn in connections:
            source_instance_id = conn.get('source')
            target_instance_id = conn.get('target')

            # Find the source and target device instances to get their names
            source_instance = find_instance_by_id(source_instance_id)
            target_instance = find_instance_by_id(target_instance_id)

            exported_data.append({
                'Source Device Name': source_instance.get('instance_name', 'Unknown') if source_instance else 'Unknown', # Get instance name, default to 'Unknown' if not found
                'Source Device ID': source_instance_id,
                'Source Port': conn.get('source_port'),
                'Target Device Name': target_instance.get('instance_name', 'Unknown') if target_instance else 'Unknown', # Get instance name, default to 'Unknown' if not found
                'Target Device ID': target_instance_id,
                'Target Port': conn.get('target_port'),
                # Add other relevant fields if needed
            })

        # Create a pandas DataFrame from the data
        df = pd.DataFrame(exported_data)

        # Create an in-memory Excel file
        output = io.BytesIO()
        # Use the ExcelWriter context manager for better handling
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
             df.to_excel(writer, index=False, sheet_name='Connections')

        # Move to the beginning of the stream
        output.seek(0)

        # Return the Excel file as a response
        return send_file(
            output,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            download_name='connections.xlsx',
            as_attachment=True
        )

    except Exception as e:
        print(f'Error exporting connections to Excel: {e}')
        # Consider returning a more informative error to the frontend
        return jsonify({'message': f'Internal server error during Excel export: {e}'}), 500

@app.route('/delete_device_type_2/<string:type_name>', methods=['DELETE'])
def delete_device_type_2(type_name):
    try:
        load_all_data_before_request()

        # Check if any device instances of this type exist
        if any(instance['device_type'] == type_name for instance in device_instances):
            return jsonify({'message': f'Cannot delete device type {type_name} because existing device instances use this type. Please delete the instances first.'}), 400

        # Find the device type by name and remove it
        global device_types
        initial_count = len(device_types)
        device_types = [device for device in device_types if device['name'] != type_name]

        if len(device_types) < initial_count:
            save_data(DEVICE_TYPES_FILE, device_types)
            print(f'Deleted device type: {type_name}')
            return jsonify({'message': 'Device type deleted successfully'}), 200
        else:
            return jsonify({'message': f'Device type {type_name} not found'}), 404

    except Exception as e:
        print(f'Error deleting device type: {e}')
        return jsonify({'message': 'Internal server error during device type deletion'}), 500

@app.route('/export_rack_panels_excel')
def export_rack_panels_excel():
    from collections import defaultdict
    import re
    
    # 按机房和排对机柜分组
    racks_by_room_row = defaultdict(lambda: defaultdict(list))
    for rack in racks:
        if rack['name'] and rack.get('room_id'):
            row = rack['name'][0].upper()
            racks_by_room_row[rack['room_id']][row].append(rack)
    
    # 对每个排内的机柜按照编号排序
    def extract_number(rack_name):
        """从机柜名称中提取数字部分用于排序"""
        match = re.search(r'(\d+)', rack_name)
        return int(match.group(1)) if match else 0
    
    # 对所有排内的机柜按编号排序
    for room_id in racks_by_room_row:
        for row in racks_by_room_row[room_id]:
            racks_by_room_row[room_id][row].sort(key=lambda rack: extract_number(rack['name']))
    
    # 机房ID到名称的映射
    room_id_to_name = {room['id']: room['name'] for room in rooms}
    sorted_rooms = sorted(room_id_to_name.keys(), key=lambda rid: room_id_to_name[rid])
    
    # 所有排的集合
    all_rows = set()
    for room_rows in racks_by_room_row.values():
        all_rows.update(room_rows.keys())
    sorted_rows = sorted(all_rows)
    
    # 按机柜ID分组设备
    devices_by_rack = defaultdict(list)
    for dev in device_instances:
        if dev.get('rack_id'):
            devices_by_rack[dev['rack_id']].append(dev)
    
    # 设备类型映射
    devtype_map = {dt['name']: dt for dt in device_types}

    wb = Workbook()
    wb.remove(wb.active)  # 删除默认sheet
    
    # 定义样式
    thin_border = Side(border_style="thin", color="000000")
    thick_border = Side(border_style="thick", color="000000")
    
    normal_border_style = Border(left=thin_border, right=thin_border, top=thin_border, bottom=thin_border)
    
    # 背景颜色
    row_title_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    u_number_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    device_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    interval_fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")
    
    # 对齐方式
    center_alignment = Alignment(horizontal="center", vertical="center")
    
    # 为每个机房创建一个sheet
    for room_id in sorted_rooms:
        room_name = room_id_to_name[room_id]
        ws = wb.create_sheet(title=room_name[:31])
        room_rows = racks_by_room_row[room_id]
        if not room_rows:
            continue
            
        current_row = 2
        current_col = 2
        
        # 遍历每个排
        for row in sorted_rows:
            racks_in_row = room_rows.get(row, [])
            if not racks_in_row:
                continue
                
            # 计算本排的最大U
            max_u = max(rack['height_u'] for rack in racks_in_row)
            n_racks = len(racks_in_row)
            
            # 每个机柜4列：左U数、设备名、右U数、间隔
            total_cols = 1 + n_racks * 4
            
            # 设置列宽
            ws.column_dimensions[get_column_letter(current_col)].width = 6  # 排名
            for i, rack in enumerate(racks_in_row):
                col_offset = current_col + 1 + i * 4
                ws.column_dimensions[get_column_letter(col_offset)].width = 4  # 左U数列
                ws.column_dimensions[get_column_letter(col_offset + 1)].width = 20  # 设备名称
                ws.column_dimensions[get_column_letter(col_offset + 2)].width = 4  # 右U数列
                ws.column_dimensions[get_column_letter(col_offset + 3)].width = 3  # 间隔
            
            # 左侧排名合并单元格
            ws.merge_cells(start_row=current_row, start_column=current_col, 
                          end_row=current_row + max_u, end_column=current_col)
            cell = ws.cell(row=current_row, column=current_col)
            cell.value = f'{row}'
            cell.border = normal_border_style
            cell.fill = row_title_fill
            cell.alignment = center_alignment
            
            # 机柜标题
            for i, rack in enumerate(racks_in_row):
                col_offset = current_col + 1 + i * 4
                
                # 左U列标
                u_cell = ws.cell(row=current_row, column=col_offset)
                u_cell.value = 'U'
                u_cell.border = normal_border_style
                u_cell.alignment = center_alignment
                
                # 机柜名标
                rack_dep = rack.get('department')
                rack_cell = ws.cell(row=current_row, column=col_offset + 1)
                rack_cell.value = rack['name'] if not rack_dep else f"{rack['name']} ({rack_dep})"
                rack_cell.border = normal_border_style
                rack_cell.alignment = center_alignment
                
                # 右U列标
                right_u_cell = ws.cell(row=current_row, column=col_offset + 2)
                right_u_cell.value = 'U'
                right_u_cell.border = normal_border_style
                right_u_cell.alignment = center_alignment
                
                # 间隔列合并
                ws.merge_cells(start_row=current_row, start_column=col_offset + 3, 
                              end_row=current_row + max_u, end_column=col_offset + 3)
                interval_cell = ws.cell(row=current_row, column=col_offset + 3)
                interval_cell.border = normal_border_style
                interval_cell.fill = interval_fill
            
            # 生成U数和设备名称
            for u_idx in range(max_u):
                # U编号到max_u，底部是1
                u_number = u_idx + 1
                row_pos = current_row + max_u - u_idx
                
                for i, rack in enumerate(racks_in_row):
                    col_offset = current_col + 1 + i * 4
                    rack_height = rack.get('height_u', max_u)
                    
                    # 写入左侧U编号
                    u_cell = ws.cell(row=row_pos, column=col_offset)
                    if u_number <= rack_height:
                        u_cell.value = u_number
                    else:
                        u_cell.value = ''
                    u_cell.border = normal_border_style
                    u_cell.fill = u_number_fill
                    u_cell.alignment = center_alignment
                    
                    # 写入右侧U编号
                    right_u_cell = ws.cell(row=row_pos, column=col_offset + 2)
                    if u_number <= rack_height:
                        right_u_cell.value = u_number
                    else:
                        right_u_cell.value = ''
                    right_u_cell.border = normal_border_style
                    right_u_cell.fill = u_number_fill
                    right_u_cell.alignment = center_alignment
                    
                    # 检查该U位是否有设备
                    rack_id = rack['id']
                    device_at_u = None
                    device_height = 0
                    
                    for dev in devices_by_rack.get(rack_id, []):
                        dev_type = devtype_map.get(dev['device_type'])
                        dev_height = dev_type['height_u'] if dev_type else 1
                        dev_bottom_u = dev['rack_u']
                        dev_top_u = dev_bottom_u + dev_height - 1
                        
                        # 检查当前U是否在设备范围内
                        if dev_bottom_u <= u_number <= dev_top_u:
                            device_at_u = dev
                            device_height = dev_height
                            break
                    
                    # 写入设备名称
                    if device_at_u and u_number == device_at_u['rack_u']:
                        # 这是设备的底部U，创建合并单元格
                        merge_start = row_pos - device_height + 1
                        merge_end = row_pos
                        if device_height > 1:
                            ws.merge_cells(start_row=merge_start, start_column=col_offset + 1, 
                                          end_row=merge_end, end_column=col_offset + 1)
                        
                        device_cell = ws.cell(row=merge_start, column=col_offset + 1)
                        device_cell.value = device_at_u['instance_name']
                        device_cell.border = normal_border_style
                        device_cell.fill = device_fill
                        device_cell.alignment = center_alignment
                        
                        # 为合并单元格的其他部分设置样式
                        if device_height > 1:
                            for r in range(merge_start, merge_end + 1):
                                cell = ws.cell(row=r, column=col_offset + 1)
                                cell.border = normal_border_style
                                cell.fill = device_fill
                                cell.alignment = center_alignment
                        
                    elif not device_at_u:
                        # 空U
                        empty_cell = ws.cell(row=row_pos, column=col_offset + 1)
                        empty_cell.value = ''
                        empty_cell.border = normal_border_style
                        empty_cell.alignment = center_alignment
            
            # 关键功能：为整个表格区域设置外边框（粗边框）
            table_start_row = current_row
            table_end_row = current_row + max_u
            table_start_col = current_col
            table_end_col = current_col + total_cols - 1
            
            # 设置外边框 只影响边界，不影响内部格
            for r in range(table_start_row, table_end_row + 1):
                for c in range(table_start_col, table_end_col + 1):
                    cell = ws.cell(row=r, column=c)
                    
                    # 保持原有边框，只在边界添加粗边框
                    left = thick_border if c == table_start_col else cell.border.left
                    right = thick_border if c == table_end_col else cell.border.right
                    top = thick_border if r == table_start_row else cell.border.top
                    bottom = thick_border if r == table_end_row else cell.border.bottom
                    
                    cell.border = Border(left=left, right=right, top=top, bottom=bottom)
            
            # 移到下一排
            current_row += max_u + 2
    
    # 保存到内存
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    
    return send_file(output, as_attachment=True, download_name='rack_panels.xlsx', 
                    mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

# --- Batch Rack Mount API for Data Consistency ---
@app.route('/batch_rack_mount', methods=['POST'])
def batch_rack_mount():
    """Atomically mounts a batch of devices to a single rack with spacing and validation."""
    try:
        load_all_data_before_request()
        data = request.get_json()
        rack_id = data.get('rack_id')
        
        # 兼容两种调用格式
        if 'device_updates' in data:
            # 单设备上架格式 兼容原有调用
            device_updates = data.get('device_updates', [])
            if not device_updates:
                return jsonify({'message': '缺少设备更新信息'}), 400
            
            device_ids = [update['device_id'] for update in device_updates]
            start_u = device_updates[0].get('rack_u', 1)
            spacing_u = 0  # 单设备模式不需要间隔
        else:
            # 批量上架格式
            device_ids = data.get('device_ids', [])
            start_u = data.get('start_u')
            spacing_u = data.get('spacing_u', 0)

        # --- Basic Validation ---
        if not all([rack_id, device_ids, start_u is not None]):
            return jsonify({'message': '缺少必要参数 (rack_id, device_ids, start_u)'}), 400
        try:
            start_u = int(start_u)
            spacing_u = int(spacing_u)
            if start_u < 1 or spacing_u < 0:
                raise ValueError()
        except (ValueError, TypeError):
            return jsonify({'message': '起始U位和间隔U位必须为有效的正整数'}), 400

        global device_instances, device_types, racks

        # --- Comprehensive Validation ---
        target_rack = find_rack_by_id(rack_id)
        if not target_rack:
            return jsonify({'message': '目标机柜未找到'}), 404
        
        # Check if all devices exist and are unmounted
        devices_to_mount = []
        for device_id in device_ids:
            instance = find_instance_by_id(device_id)
            if not instance:
                return jsonify({'message': f'设备 ID {device_id} 未找到'}), 404
            if instance.get('rack_id'):
                return jsonify({'message': f'设备 "{instance.get("instance_name")}" 已上架，请勿重复操作'}), 400
            devices_to_mount.append(instance)

        # Get occupied U-slots in the target rack
        occupied_slots = get_occupied_slots(target_rack['id'])

        # Calculate required slots and check for conflicts
        required_slots = {} # U -> device_name
        current_u = start_u
        total_height_needed = 0

        for device in devices_to_mount:
            device_type_info = find_type_by_name(device.get('device_type'))
            height = device_type_info.get('height_u', 1) if device_type_info else 1
            
            # Check for conflicts for each U the device will occupy
            for i in range(current_u, current_u + height):
                if i in occupied_slots:
                    return jsonify({'message': f'操作失败：U{i} 已被设备 "{occupied_slots[i]}" 占用'}), 409
                if i in required_slots: # Should not happen, but a safeguard
                    return jsonify({'message': f'操作失败：计算出的U{i} 内部冲突'}), 500
                required_slots[i] = device.get('instance_name')

            total_height_needed += height
            current_u += height + spacing_u # Add spacing for the next device
        
        end_u = current_u - spacing_u - 1 # The U of the top of the last device
        if end_u > target_rack.get('height_u', 42):
            return jsonify({'message': f'机柜容量不足。需要U{end_u}，但机柜总高{target_rack.get("height_u", 42)}U'}), 409

        # --- Perform Mount Operation ---
        current_u = start_u
        for device in devices_to_mount:
            device_type_info = find_type_by_name(device.get('device_type'))
            height = device_type_info.get('height_u', 1) if device_type_info else 1
            
            # Update device instance
            instance_to_update = find_instance_by_id(device['id'])
            instance_to_update['rack_id'] = rack_id
            instance_to_update['rack_u'] = current_u

            # Add device to rack's list
            if 'devices' not in target_rack:
                target_rack['devices'] = []
            if not any(d.get('instance_id') == device['id'] for d in target_rack['devices']):
                target_rack['devices'].append({
                    'instance_id': device['id'],
                    'rack_u': current_u
                })

            current_u += height + spacing_u

        save_data(DEVICE_INSTANCES_FILE, device_instances)
        save_data(RACKS_FILE, racks)

        return jsonify({'message': f'成功上架 {len(devices_to_mount)} 台设备至机柜 "{target_rack.get("name")}"'}), 200

    except Exception as e:
        print(f"Error during batch rack mount: {e}")
        return jsonify({'message': '批量上架期间发生内部错误'}), 500

@app.route('/devices/batch_delete', methods=['POST'])
def batch_delete_devices():
    """Atomically deletes a batch of devices."""
    try:
        load_all_data_before_request()
        data = request.get_json()
        device_ids_to_delete = data.get('device_ids', [])

        if not isinstance(device_ids_to_delete, list) or not device_ids_to_delete:
            return jsonify({'message': 'Please provide a list of device IDs to delete.'}), 400

        global device_instances, connections, racks
        
        # --- Validation ---
        error_details = []
        for device_id in device_ids_to_delete:
            instance = find_instance_by_id(device_id)
            if not instance:
                # Might have been deleted by another process, but we can ignore it.
                continue
            if instance.get('rack_id'):
                error_details.append(f'设备 "{instance.get("instance_name")}" 必须先下架才能删除')
        
        if error_details:
            return jsonify({
                'message': '验证失败，部分设备仍处于上架状态',
                'errors': error_details
            }), 400

        # --- Perform Deletion in a single operation ---
        
        original_instance_count = len(device_instances)
        ids_to_delete_set = set(device_ids_to_delete)
        
        # Filter out devices to be deleted
        device_instances[:] = [inst for inst in device_instances if inst['id'] not in ids_to_delete_set]
        
        # Find connections associated with the deleted devices to free up ports on the other end
        ports_to_free = []
        remaining_connections = []
        for conn in connections:
            source_deleted = conn['source'] in ids_to_delete_set
            target_deleted = conn['target'] in ids_to_delete_set
            
            if source_deleted and not target_deleted:
                ports_to_free.append({'instance_id': conn['target'], 'port_name': conn['target_port']})
            elif not source_deleted and target_deleted:
                ports_to_free.append({'instance_id': conn['source'], 'port_name': conn['source_port']})
            
            if not source_deleted and not target_deleted:
                remaining_connections.append(conn)

        connections[:] = remaining_connections
        
        # Free up the ports on the remaining connected devices
        for port_info in ports_to_free:
            update_interface_status(port_info['instance_id'], port_info['port_name'], 'available')

        # Since validation ensures devices are unmounted, we don't need to check racks.
        # However, a safety check is good practice.
        for rack in racks:
            if 'devices' in rack:
                rack['devices'][:] = [dev_id for dev_id in rack['devices'] if dev_id not in ids_to_delete_set]

        # Save all changes once
        save_data(DEVICE_INSTANCES_FILE, device_instances)
        save_data(CONNECTIONS_FILE, connections)
        save_data(RACKS_FILE, racks)
        
        deleted_count = original_instance_count - len(device_instances)

        return jsonify({'message': f'成功删除 {deleted_count} 个设备'}), 200

    except Exception as e:
        print(f"Error during batch device deletion: {e}")
        return jsonify({'message': '批量删除期间发生内部错误'}), 500

# --- Routes for Rooms --- #

@app.route('/add_room', methods=['POST'])
def add_room():
    try:
        data = request.get_json()
        if not data or 'name' not in data:
            return jsonify({'message': 'Invalid data provided. \'name\' is required.'}), 400

        room_name = data['name'].strip()
        if not room_name:
             return jsonify({'message': 'Room name cannot be empty.'}), 400

        # Check if room already exists (case-insensitive)
        if any(room['name'].lower() == room_name.lower() for room in rooms):
            return jsonify({'message': f'Room "{room_name}" already exists.'}), 409 # Conflict

        # Generate a simple ID for the room (e.g., timestamp)
        room_id = f'room_{int(time.time() * 1000)}'

        new_room = {
            'id': room_id,
            'name': room_name,
            # Add other room properties here if needed (e.g., description, location)
        }
        rooms.append(new_room)

        save_data(ROOMS_FILE, rooms)

        print(f"[API] add_room: Added new room '{room_name}' ({room_id})")
        return jsonify({'message': 'Room added successfully', 'room': new_room}), 201

    except Exception as e:
        print(f'Error adding room: {e}')
        return jsonify({'message': 'Internal server error during room addition'}), 500

@app.route('/get_rooms', methods=['GET'])
def get_rooms():
    try:
        # Ensure latest data is loaded before returning (redundant with before_request, but safe)
        load_all_data_before_request()
        return jsonify({'rooms': rooms}), 200
    except Exception as e:
        print(f'Error getting rooms: {e}')
        return jsonify({'message': 'Internal server error during getting rooms'}), 500

@app.route('/delete_room/<string:room_id>', methods=['DELETE'])
def delete_room(room_id):
    try:
        load_all_data_before_request()

        # Check if any racks are in this room
        global racks
        racks_in_room = [rack for rack in racks if rack['room_id'] == room_id]
        if racks_in_room:
             # If there are racks, delete them first (this will also handle deleting devices in those racks)
             for rack_to_delete in racks_in_room:
                 _internal_delete_rack(rack_to_delete['id']) # Call the internal logic function

        # Find the room by ID and remove it
        global rooms
        initial_count = len(rooms)
        rooms = [room for room in rooms if room['id'] != room_id]

        if len(rooms) < initial_count:
            save_data(ROOMS_FILE, rooms)
            print(f'[API] delete_room: Deleted room with ID {room_id}')
            return jsonify({'message': 'Room deleted successfully'}), 200
        else:
            return jsonify({'message': f'Room with ID {room_id} not found'}), 404

    except Exception as e:
        print(f'Error deleting room: {e}')
        return jsonify({'message': 'Internal server error during room deletion'}), 500

# --- Routes for Racks --- #

@app.route('/add_rack', methods=['POST'])
def add_rack():
    try:
        data = request.get_json()
        name = data.get('name', '').strip()
        room_id = data.get('room_id')
        height_u = data.get('height_u')
        department = data.get('department', '').strip()

        # Validate required fields
        if not name or not room_id or not height_u:
            return jsonify({'message': '机柜名称、所属机房和高度为必填项'}), 400

        # Validate height_u is a positive integer
        try:
            height_u = int(height_u)
            if height_u <= 0 or height_u > 100:
                return jsonify({'message': '机柜高度必须为1-100之间的正整数'}), 400
        except ValueError:
            return jsonify({'message': '机柜高度必须是正整数'}), 400

        # Check if room exists
        room_exists = False
        for room in rooms:
            if room['id'] == room_id:
                room_exists = True
                break
        if not room_exists:
            return jsonify({'message': '指定的机房不存在'}), 400

        # Check if rack name already exists in the room
        for rack in racks:
            if rack['room_id'] == room_id and rack['name'] == name:
                return jsonify({'message': f'机房中已存在名为 {name} 的机柜'}), 400

        # Generate a new unique ID
        new_id = str(uuid.uuid4())

        # Create new rack
        new_rack = {
            'id': new_id,
            'name': name,
            'room_id': room_id,
            'height_u': height_u,
            'department': department
        }

        # Add to racks list and save to file
        racks.append(new_rack)
        save_data(RACKS_FILE, racks)

        return jsonify({'message': '机柜添加成功', 'rack': new_rack})

    except Exception as e:
        print(f"Error adding rack: {str(e)}")
        return jsonify({'message': '添加机柜失败'}), 500

@app.route('/get_racks', methods=['GET'])
def get_racks():
    try:
        load_all_data_before_request()
        return jsonify({'racks': racks}), 200
    except Exception as e:
        print(f'Error getting racks: {e}')
        return jsonify({'message': 'Internal server error during getting racks'}), 500

@app.route('/batch_add_racks', methods=['POST'])
def batch_add_racks():
    """Batch add multiple racks to a room"""
    try:
        load_all_data_before_request()
        data = request.get_json()
        
        room_id = data.get('room_id')
        # 兼容前端发送的 rack_row 字段，也支持 rack_prefix
        rack_prefix = data.get('rack_prefix') or data.get('rack_row', '').strip()
        start_number = data.get('start_number')
        end_number = data.get('end_number')
        height_u = data.get('height_u')
        # department 字段是可选的，如果没有就设为空字符串
        department = data.get('department', '').strip()

        # Validate required fields - department 不是必需项
        if not all([room_id, rack_prefix, isinstance(start_number, int), isinstance(end_number, int), isinstance(height_u, int)]):
            return jsonify({'message': '缺少必要字段。收到的数据: room_id={room_id}, rack_prefix={rack_prefix}, start_number={start_number}, end_number={end_number}, height_u={height_u}'}), 400

        if start_number > end_number:
            return jsonify({'message': '起始编号不能大于结束编号'}), 400

        if end_number - start_number + 1 > 100:  # Limit batch size
            return jsonify({'message': '单次批量添加的机柜数量不能超过100'}), 400

        if height_u <= 0 or height_u > 100:
            return jsonify({'message': '机柜高度必须为1-100之间的正整数'}), 400

        # Check if room exists
        room_exists = any(room['id'] == room_id for room in rooms)
        if not room_exists:
            return jsonify({'message': '指定的机房不存在'}), 400

        # Check for existing rack names in the room
        existing_rack_names = {rack['name'].lower() for rack in racks if rack['room_id'] == room_id}
        new_racks = []
        skipped_names = []

        for i in range(start_number, end_number + 1):
            rack_name = f"{rack_prefix}{i:02d}"
            if rack_name.lower() in existing_rack_names:
                skipped_names.append(rack_name)
                continue

            new_rack = {
                'id': str(uuid.uuid4()),
                'name': rack_name,
                'room_id': room_id,
                'height_u': height_u,
                'department': department,
                'devices': []
            }
            new_racks.append(new_rack)

        if not new_racks:
            return jsonify({'message': f'未能添加任何新机柜。以下名称均已存在 {", ".join(skipped_names)}'}), 409

        # Add all new racks
        racks.extend(new_racks)
        save_data(RACKS_FILE, racks)

        # 获取房间名称用于返回消息
        room_name = next((room['name'] for room in rooms if room['id'] == room_id), '未知机房')
        created_rack_names = [rack['name'] for rack in new_racks]

        message = f'成功添加 {len(new_racks)} 个机柜'
        if skipped_names:
            message += f'，跳过 {len(skipped_names)} 个已存在的机柜 {", ".join(skipped_names)}'

        return jsonify({
            'success': True,
            'message': message,
            'room_name': room_name,
            'created_count': len(new_racks),
            'created_racks': created_rack_names,
            'duplicates': skipped_names,
            'added_count': len(new_racks),
            'skipped_count': len(skipped_names),
            'added_racks': new_racks
        }), 201

    except Exception as e:
        print(f"Error in batch add racks: {str(e)}")
        return jsonify({'message': '批量添加机柜时发生内部错误'}), 500

@app.route('/delete_rack/<string:rack_id>', methods=['DELETE'])
def delete_rack(rack_id):
    try:
        load_all_data_before_request()
        success, updated_instances_data = _internal_delete_rack(rack_id)

        if success:
            return jsonify({'message': 'Rack deleted successfully', 'updated_instances': updated_instances_data}), 200
        else:
             return jsonify({'message': f'Rack with ID {rack_id} not found'}), 404

    except Exception as e:
        print(f'Error deleting rack: {e}')
        return jsonify({'message': 'Internal server error during rack deletion'}), 500

@app.route('/batch_delete_racks', methods=['POST'])
def batch_delete_racks():
    """Batch delete multiple racks"""
    try:
        load_all_data_before_request()
        data = request.get_json()
        rack_ids = data.get('rack_ids', [])
        
        if not isinstance(rack_ids, list) or not rack_ids:
            return jsonify({'message': '请提供要删除的机柜ID列表'}), 400
        
        global racks
        deleted_count = 0
        error_details = []
        updated_instances_all = []
        
        for rack_id in rack_ids:
            # Check if rack has devices
            devices_in_rack = [inst for inst in device_instances if inst.get('rack_id') == rack_id]
            if devices_in_rack:
                rack = find_rack_by_id(rack_id)
                rack_name = rack.get('name', rack_id) if rack else rack_id
                error_details.append(f'机柜 "{rack_name}" 中还有设备，无法删除')
                continue
            
            success, updated_instances = _internal_delete_rack(rack_id)
            if success:
                deleted_count += 1
                updated_instances_all.extend(updated_instances)
        
        message = f'成功删除 {deleted_count} 个机柜'
        if error_details:
            message += f'，{len(error_details)} 个机柜删除失败'
        
        return jsonify({
            'message': message,
            'deleted_count': deleted_count,
            'errors': error_details,
            'updated_instances': updated_instances_all
        }), 200 if not error_details else 206
        
    except Exception as e:
        print(f"Error in batch delete racks: {str(e)}")
        return jsonify({'message': '批量删除机柜时发生内部错误'}), 500

@app.route('/update_rack/<string:rack_id>', methods=['PUT'])
def update_rack(rack_id):
    """Update rack information"""
    try:
        load_all_data_before_request()
        data = request.get_json()
        
        rack = find_rack_by_id(rack_id)
        if not rack:
            return jsonify({'message': '机柜未找到'}), 404
        
        # Update rack properties
        if 'name' in data:
            new_name = data['name'].strip()
            if not new_name:
                return jsonify({'message': '机柜名称不能为空'}), 400
            
            # Check for duplicate names in the same room
            for r in racks:
                if r['id'] != rack_id and r['room_id'] == rack['room_id'] and r['name'] == new_name:
                    return jsonify({'message': f'机房中已存在名为 "{new_name}" 的机柜'}), 400
            
            rack['name'] = new_name
        
        if 'department' in data:
            rack['department'] = data['department'].strip()
        
        if 'height_u' in data:
            try:
                new_height = int(data['height_u'])
                if new_height <= 0 or new_height > 100:
                    return jsonify({'message': '机柜高度必须为1-100之间的正整数'}), 400
                
                # Check if reducing height would conflict with existing devices
                devices_in_rack = [inst for inst in device_instances if inst.get('rack_id') == rack_id]
                max_used_u = 0
                for device in devices_in_rack:
                    if device.get('rack_u'):
                        device_type = find_type_by_name(device.get('device_type'))
                        device_height = device_type.get('height_u', 1) if device_type else 1
                        top_u = device['rack_u'] + device_height - 1
                        max_used_u = max(max_used_u, top_u)
                
                if new_height < max_used_u:
                    return jsonify({'message': f'无法将机柜高度设置为{new_height}U，已有设备占用到{max_used_u}U'}), 400
                
                rack['height_u'] = new_height
            except (ValueError, TypeError):
                return jsonify({'message': '机柜高度必须是正整数'}), 400
        
        save_data(RACKS_FILE, racks)
        return jsonify({'message': '机柜信息更新成功', 'rack': rack}), 200
        
    except Exception as e:
        print(f"Error updating rack: {str(e)}")
        return jsonify({'message': '更新机柜信息时发生内部错误'}), 500

@app.route('/batch_rack_unmount', methods=['POST'])
def batch_rack_unmount():
    """Batch unmount devices from racks"""
    try:
        load_all_data_before_request()
        data = request.get_json()
        device_ids = data.get('device_ids', [])
        
        if not isinstance(device_ids, list) or not device_ids:
            return jsonify({'message': '请提供要下架的设备ID列表'}), 400
        
        global device_instances, racks
        unmounted_count = 0
        error_details = []
        
        for device_id in device_ids:
            instance = find_instance_by_id(device_id)
            if not instance:
                error_details.append(f'设备 {device_id} 未找到')
                continue
            
            if not instance.get('rack_id'):
                error_details.append(f'设备 "{instance.get("instance_name")}" 未上架')
                continue
            
            # Check if device is powered on
            if instance.get('power_status') == 'on':
                error_details.append(f'设备 "{instance.get("instance_name")}" 必须先关机才能下架')
                continue
            
            # Remove from rack
            rack_id = instance['rack_id']
            rack = find_rack_by_id(rack_id)
            if rack and 'devices' in rack and device_id in rack['devices']:
                rack['devices'].remove(device_id)
            
            # Clear rack info from device
            instance['rack_id'] = None
            instance['rack_u'] = None
            unmounted_count += 1
        
        save_data(DEVICE_INSTANCES_FILE, device_instances)
        save_data(RACKS_FILE, racks)
        
        message = f'成功下架 {unmounted_count} 台设备'
        if error_details:
            message += f'，{len(error_details)} 台设备下架失败'
        
        return jsonify({
            'message': message,
            'unmounted_count': unmounted_count,
            'errors': error_details
        }), 200 if not error_details else 206
        
    except Exception as e:
        print(f"Error in batch rack unmount: {str(e)}")
        return jsonify({'message': '批量下架时发生内部错误'}), 500

@app.route('/multi_rack_batch_mount', methods=['POST'])
def multi_rack_batch_mount():
    """Mount multiple devices to multiple racks in a single operation."""
    try:
        data = request.get_json()
        device_ids = data.get('device_ids', [])
        rack_assignments = data.get('rack_assignments', [])
        
        if not device_ids or not rack_assignments:
            return jsonify({'message': '缺少设备ID列表或机柜分配信息'}), 400
        
        global device_instances, device_types, racks
        
        # --- 验证所有设备存在且未上架---
        devices_to_mount = []
        for device_id in device_ids:
            instance = find_instance_by_id(device_id)
            if not instance:
                return jsonify({'message': f'设备 ID {device_id} 未找到'}), 404
            if instance.get('rack_id'):
                return jsonify({'message': f'设备 "{instance.get("instance_name")}" 已上架，请勿重复操作'}), 400
            devices_to_mount.append(instance)
        
        # --- 验证所有机柜存在---
        target_racks = {}
        for assignment in rack_assignments:
            rack_id = assignment.get('rack_id')
            rack = find_rack_by_id(rack_id)
            if not rack:
                return jsonify({'message': f'机柜 ID {rack_id} 未找到'}), 404
            target_racks[rack_id] = rack
        
        # --- 计算每个机柜的占用情况和容量检查---
        mount_plan = []
        device_index = 0
        
        for assignment in rack_assignments:
            rack_id = assignment.get('rack_id')
            start_u = assignment.get('start_u', 1)
            device_count = assignment.get('device_count', 1)
            spacing_u = assignment.get('spacing_u', 0)
            
            try:
                start_u = int(start_u)
                device_count = int(device_count)
                spacing_u = int(spacing_u)
            except (ValueError, TypeError):
                return jsonify({'message': f'机柜 {rack_id} 的参数无效'}), 400
            
            # 检查是否有足够的设备
            if device_index + device_count > len(devices_to_mount):
                return jsonify({'message': f'设备数量不足，机柜 {target_racks[rack_id].get("name")} 需要 {device_count} 台设备'}), 400
            
            # 获取该机柜的占用情况
            occupied_slots = get_occupied_slots(rack_id)
            current_u = start_u
            
            # 计算每个设备的位置
            for i in range(device_count):
                device = devices_to_mount[device_index + i]
                device_type_name = device.get('device_type')
                device_type_info = next((dt for dt in device_types if dt['name'] == device_type_name), None)
                height = device_type_info.get('height_u', 1) if device_type_info else 1
                
                # 检查冲突
                for i in range(current_u, current_u + height):
                    if i in occupied_slots:
                        return jsonify({'message': f'机柜 "{target_racks[rack_id].get("name")}" U{i} 已被占用'}), 409
                    if i > target_racks[rack_id].get('height_u', 42):
                        return jsonify({'message': f'机柜 "{target_racks[rack_id].get("name")}" 容量不足，需要U{i}但机柜总高{target_racks[rack_id].get("height_u", 42)}U'}), 409
                
                # 添加到挂载计划
                mount_plan.append({
                    'device': device,
                    'rack_id': rack_id,
                    'rack_u': current_u
                })
                
                # 更新下一个设备的起始U
                current_u += height + spacing_u
            
            device_index += device_count
        
        # 执行挂载计划
        mounted_count = 0
        for plan in mount_plan:
            instance_to_update = plan['device']
            rack_id = plan['rack_id']
            rack_u = plan['rack_u']
            
            # 更新设备实例
            instance_to_update['rack_id'] = rack_id
            instance_to_update['rack_u'] = rack_u
            
            # 添加设备到机柜列表
            rack = target_racks[rack_id]
            if 'devices' not in rack:
                rack['devices'] = []
            rack['devices'].append({
                'instance_id': instance_to_update['id'],
                'rack_u': rack_u
            })
            
            mounted_count += 1
        
        # 保存更改
        save_data(DEVICE_INSTANCES_FILE, device_instances)
        save_data(RACKS_FILE, racks)
        
        # 生成结果信息
        rack_names = [target_racks[assignment['rack_id']].get('name', assignment['rack_id']) for assignment in rack_assignments]
        message = f'成功上架 {mounted_count} 台设备到 {len(rack_assignments)} 个机柜 {", ".join(rack_names)}'
        
        return jsonify({
            'message': message,
            'mounted_count': mounted_count,
            'rack_assignments': [{
                'rack_id': rack_id,
                'rack_name': target_racks[rack_id].get('name', rack_id),
                'device_count': sum(1 for plan in mount_plan if plan['rack_id'] == rack_id)
            } for rack_id in target_racks]
        }), 200
        
    except Exception as e:
        print(f"Error in multi-rack batch mount: {str(e)}")
        return jsonify({'message': '批量上架时发生内部错误'}), 500

# Helper function for the core logic of deleting a rack
def _internal_delete_rack(rack_id):
    """
    Internal logic to delete a rack and its associated devices.
    This function contains the business logic and should not be a view function.
    Returns: (bool, list) -> (success_status, list_of_updated_instances)
    """
    global device_instances, racks
    devices_in_rack = [instance for instance in device_instances if instance.get('rack_id') == rack_id]
    updated_instances_during_cascade = []

    if devices_in_rack:
        for instance_to_delete in devices_in_rack:
            success, updated_instances_data = perform_delete_device_instance(instance_to_delete['id'])
            if success:
                updated_instances_during_cascade.extend(updated_instances_data)

    initial_count = len(racks)
    racks = [rack for rack in racks if rack['id'] != rack_id]

    if len(racks) < initial_count:
        save_data(RACKS_FILE, racks)
        print(f'[INTERNAL] Deleted rack {rack_id} and its {len(devices_in_rack)} devices.')
        return True, updated_instances_during_cascade
    else:
        return False, []

@app.route('/device_management')
def device_management():
    return render_template('device_management.html')

@app.route('/update_device_instance/<string:instance_id>', methods=['PUT'])
def update_device_instance(instance_id):
    try:
        data = request.get_json()
        new_name = data.get('instance_name', '').strip()

        if not new_name:
            return jsonify({'message': '设备实例名称不能为空'}), 400

        # Check for duplicate names
        for instance in device_instances:
            if instance['id'] != instance_id and instance['instance_name'] == new_name:
                return jsonify({'message': f'已存在名称为 "{new_name}" 的设备实例'}), 400

        # Find and update the instance
        instance_found = False
        for instance in device_instances:
            if instance['id'] == instance_id:
                instance['instance_name'] = new_name
                instance_found = True
                break

        if not instance_found:
            return jsonify({'message': '设备实例未找到'}), 404

        save_data(DEVICE_INSTANCES_FILE, device_instances)
        return jsonify({'message': '设备实例更新成功'}), 200

    except Exception as e:
        print(f"Error updating device instance: {str(e)}")
        return jsonify({'message': '更新设备实例时发生内部错误'}), 500

@app.route('/devices/batch_power', methods=['POST'])
def batch_power_devices():
    """Batch power on or off devices"""
    try:
        load_all_data_before_request()
        data = request.get_json()
        device_ids = data.get('device_ids', [])
        status = data.get('status') # 'on' or 'off'

        if not device_ids or status not in ['on', 'off']:
            return jsonify({'message': '缺少 device_ids 或无效的状态'}), 400

        success_count = 0
        error_details = []
        
        for device_id in device_ids:
            instance = find_instance_by_id(device_id)
            if not instance:
                error_details.append(f"设备 {device_id} 未找到")
                continue

            # Powering on requires the device to be mounted
            if status == 'on' and not instance.get('rack_id'):
                error_details.append(f"设备 {instance.get('instance_name')} 未上架，无法上电")
                continue
            
            instance['power_status'] = status
            success_count += 1
        
        save_data(DEVICE_INSTANCES_FILE, device_instances)

        message = f"成功 {status_map.get(status, '')} {success_count} 个设备"
        if error_details:
             message += f" {len(error_details)} 个设备操作失败"

        return jsonify({
            'message': message,
            'success_count': success_count,
            'errors': error_details
        }), 200 if not error_details else 206

    except Exception as e:
        print(f"Error in batch power operation: {str(e)}")
        return jsonify({'message': '批量电源操作时发生内部错误'}), 500

@app.route('/import_device_instances', methods=['POST'])
def import_device_instances():
    """Import device instances from external data."""
    try:
        data = request.get_json()
        if not data or 'devices' not in data:
            return jsonify({'message': '无效的导入数据'}), 400

        devices_to_import = data['devices']
        if not isinstance(devices_to_import, list):
            return jsonify({'message': '设备数据必须是列表格式'}), 400

        load_all_data_before_request()

        imported_devices = []
        skipped_devices = []
        error_devices = []

        for device in devices_to_import:
            device_name = device.get('name', '').strip()
            device_type = device.get('type', '').strip()

            if not device_name or not device_type:
                error_devices.append({
                    'name': device_name or '未知',
                    'type': device_type or '未知',
                    'reason': '设备名称或型号为空'
                })
                continue

            # Check for duplicate names
            if any(inst['instance_name'].lower() == device_name.lower() for inst in device_instances):
                skipped_devices.append({
                    'name': device_name,
                    'type': device_type,
                    'reason': '设备名称已存在'
                })
                continue

            # Check if device type exists
            device_type_definition = next((dt for dt in device_types if dt['name'] == device_type), None)
            if not device_type_definition:
                error_devices.append({
                    'name': device_name,
                    'type': device_type,
                    'reason': '设备型号未定义'
                })
                continue

            # Create new device instance
            instance_id = str(uuid.uuid4())
            new_instance = {
                'id': instance_id,
                'instance_name': device_name,
                'device_type': device_type,
                'interfaces': device_type_definition.get('interfaces', []),
                'power_status': 'off'
            }

            imported_devices.append(new_instance)

        # Add all valid devices
        if imported_devices:
            device_instances.extend(imported_devices)
            save_data(DEVICE_INSTANCES_FILE, device_instances)

        # Generate response message
        message_parts = []
        if imported_devices:
            message_parts.append(f"成功导入 {len(imported_devices)} 个设备")
        if skipped_devices:
            message_parts.append(f"跳过 {len(skipped_devices)} 个重复设备")
        if error_devices:
            message_parts.append(f"无法导入 {len(error_devices)} 个设备")

        message = "; ".join(message_parts) if message_parts else "没有设备被导入"

        return jsonify({
            'message': message,
            'imported': len(imported_devices),
            'skipped': len(skipped_devices),
            'errors': len(error_devices),
            'error_details': error_devices,
            'skipped_details': skipped_devices
        }), 200 if imported_devices else 400

    except Exception as e:
        print(f"Error importing device instances: {str(e)}")
        return jsonify({'message': '导入设备时发生内部错误'}), 500

# Route to get device instances for import validation

@app.route('/connection_tool')
def connection_tool():
    """连线工具页面"""
    return render_template('connection_tool.html')

@app.route('/get_connections', methods=['GET'])
def get_connections():
    """获取所有连接"""
    try:
        load_all_data_before_request()
        # 直接返回全局变量中的连接列表
        return jsonify(connections), 200
    except Exception as e:
        print(f"Error getting connections: {str(e)}")
        return jsonify({'message': '获取连接时发生错误'}), 500

@app.route('/add_connection', methods=['POST'])
def add_connection():
    """添加连接"""
    try:
        # 确保数据已加载
        load_all_data_before_request()
        
        data = request.get_json()
        print(f"Received connection data: {data}")
        
        if not data:
            print("No data received")
            return jsonify({'message': '没有收到数据'}), 400
            
        # 检查必要字段
        required_fields = ['source_device', 'target_device', 'source_port', 'target_port']
        missing_fields = [field for field in required_fields if not data.get(field)]
        if missing_fields:
            print(f"Missing fields: {missing_fields}")
            return jsonify({'message': f'缺少必要字段: {", ".join(missing_fields)}'}), 400

        # 检查设备是否存在
        source_device = find_instance_by_id(data['source_device'])
        target_device = find_instance_by_id(data['target_device'])

        if not source_device:
            print(f"Source device not found: {data['source_device']}")
            return jsonify({'message': f'源设备不存在: {data["source_device"]}'}), 404
        if not target_device:
            print(f"Target device not found: {data['target_device']}")
            return jsonify({'message': f'目标设备不存在: {data["target_device"]}'}), 404

        # 检查接口是否存在
        source_interface = next((iface for iface in source_device.get('interfaces', []) 
                               if iface['name'] == data['source_port']), None)
        target_interface = next((iface for iface in target_device.get('interfaces', [])
                               if iface['name'] == data['target_port']), None)

        if not source_interface:
            print(f"Source interface not found: {data['source_port']}")
            return jsonify({'message': f'源接口不存在: {data["source_port"]}'}), 404
        if not target_interface:
            print(f"Target interface not found: {data['target_port']}")
            return jsonify({'message': f'目标接口不存在: {data["target_port"]}'}), 404

        # 检查接口是否已被占用
        if source_interface.get('connected_to'):
            print(f"Source interface already connected: {data['source_port']}")
            return jsonify({'message': f'源接口已被占用: {data["source_port"]}'}), 400
        if target_interface.get('connected_to'):
            print(f"Target interface already connected: {data['target_port']}")
            return jsonify({'message': f'目标接口已被占用: {data["target_port"]}'}), 400

        # 生成连接ID
        connection_id = str(uuid.uuid4())

        # 创建连接记录
        connection = {
            'id': connection_id,
            'source_device': data['source_device'],
            'target_device': data['target_device'],
            'source_port': data['source_port'],
            'target_port': data['target_port']
        }

        # 更新接口连接状态
        source_interface['connected_to'] = {
            'device_id': data['target_device'],
            'port_name': data['target_port'],
            'connection_id': connection_id
        }
        target_interface['connected_to'] = {
            'device_id': data['source_device'],
            'port_name': data['source_port'],
            'connection_id': connection_id
        }

        # 添加到连接列表
        global connections
        connections.append(connection)

        # 保存更改
        print(f"Saving connection data: {connection}")
        save_data(DEVICE_INSTANCES_FILE, device_instances)
        save_data(CONNECTIONS_FILE, connections)

        return jsonify({
            'message': '连接创建成功',
            'connection': connection
        }), 201

    except Exception as e:
        print(f"Error creating connection: {str(e)}")
        import traceback
        print(f"Traceback: {traceback.format_exc()}")
        return jsonify({'message': f'创建连接时发生错误: {str(e)}'}), 500

@app.route('/delete_connection/<connection_id>', methods=['DELETE'])
def delete_connection(connection_id):
    """删除连接"""
    try:
        load_all_data_before_request()
        
        # 查找连接
        global connections
        connection = None
        for conn in connections:
            if conn.get('id') == connection_id:
                connection = conn
                break

        if not connection:
            return jsonify({
                'message': '指定的连接不存在'
            }), 404

        # 查找相关设备和接口
        source_device = None
        target_device = None
        for device in device_instances:
            if device['id'] == connection['source_device']:
                source_device = device
            if device['id'] == connection['target_device']:
                target_device = device
            if source_device and target_device:
                break

        if source_device and target_device:
            # 清除接口连接状态
            for iface in source_device.get('interfaces', []):
                if iface['name'] == connection['source_port']:
                    iface['connected_to'] = None
                    break
            for iface in target_device.get('interfaces', []):
                if iface['name'] == connection['target_port']:
                    iface['connected_to'] = None
                    break

        # 删除连接记录
        connections.remove(connection)

        # 保存更改
        save_data(DEVICE_INSTANCES_FILE, device_instances)
        save_data(CONNECTIONS_FILE, connections)

        return jsonify({
            'message': '连接删除成功'
        }), 200

    except Exception as e:
        print(f"Error deleting connection: {str(e)}")
        return jsonify({'message': '删除连接时发生错误'}), 500

# 设备组管理API
@app.route('/get_device_groups', methods=['GET'])
def get_device_groups():
    """获取所有设备组"""
    try:
        return jsonify({"groups": device_groups})
    except Exception as e:
        app.logger.error(f"Error getting device groups: {str(e)}")
        return jsonify({"error": "Failed to get device groups"}), 500

@app.route('/get_device_group/<group_id>', methods=['GET'])
def get_device_group(group_id):
    """获取特定设备组的信息"""
    group = next((g for g in device_groups if g['id'] == group_id), None)
    if not group:
        return jsonify({"error": "Group not found"}), 404
    return jsonify(group)

@app.route('/delete_device_group/<group_id>', methods=['DELETE'])
def delete_device_group(group_id):
    """删除设备组"""
    try:
        # 找到要删除的组
        group = next((g for g in device_groups if g['id'] == group_id), None)
        if not group:
            return jsonify({"error": "Group not found"}), 404

        # 删除组内的所有连接
        connections[:] = [c for c in connections if c.get('group_id') != group_id]
        save_data(CONNECTIONS_FILE, connections)

        # 删除组
        device_groups[:] = [g for g in device_groups if g['id'] != group_id]
        save_data(DEVICE_GROUPS_FILE, device_groups)

        return jsonify({"message": "Device group deleted successfully"}), 200

    except Exception as e:
        app.logger.error(f"Error deleting device group: {str(e)}")
        return jsonify({"error": "Internal server error"}), 500

def generate_group_id(group_type, index):
    """
    生成设备组ID
    :param group_type: 组类型 (stack/mlag)
    :param index: 序号
    :return: 组ID
    """
    return f"{group_type}_group_{str(index).zfill(3)}"

def get_next_group_index(group_type):
    """
    获取下一个可用的组序号
    :param group_type: 组类型 (stack/mlag)
    :return: 序号
    """
    existing_groups = [g for g in device_groups if g.get('type') == group_type]
    if not existing_groups:
        return 1
    
    existing_indices = []
    for group in existing_groups:
        try:
            index = int(group['id'].split('_')[-1])
            existing_indices.append(index)
        except (ValueError, IndexError):
            continue
    
    if not existing_indices:
        return 1
    
    return max(existing_indices) + 1

def create_group_connections(group_id, device1_id, device2_id, template_config, group_type):
    """
    创建设备组内的连接关系
    :param group_id: 组ID
    :param device1_id: 设备1 ID
    :param device2_id: 设备2 ID
    :param template_config: 模板配置
    :param group_type: 组类型 (stack/mlag)
    :return: 连接列表
    """
    new_connections = []
    
    # 创建设备互联连接
    device_links = template_config['device_links']
    
    # 创建peer-link连接
    for i, port in enumerate(device_links['peer_link']['ports']):
        connection = {
            "id": str(uuid.uuid4()),
            "source_device": device1_id,
            "target_device": device2_id,
            "source_port": port,
            "target_port": port,  # 使用相同端口号
            "type": "peer-link",
            "group_id": group_id,
            "speed": device_links['peer_link']['speed']
        }
        new_connections.append(connection)
    
    # 创建heartbeat连接
    for i, port in enumerate(device_links['heartbeat']['ports']):
        connection = {
            "id": str(uuid.uuid4()),
            "source_device": device1_id,
            "target_device": device2_id,
            "source_port": port,
            "target_port": port,  # 使用相同端口号
            "type": "heartbeat",
            "group_id": group_id,
            "speed": device_links['heartbeat']['speed']
        }
        new_connections.append(connection)
    
    # 创建服务端口连接
    service_ports = template_config['service_ports']
    
    # 创建上行端口连接
    for port in service_ports['uplink']['ports']:
        connection = {
            "id": str(uuid.uuid4()),
            "source_device": device1_id,
            "target_device": device2_id,
            "source_port": port,
            "target_port": port,  # 使用相同端口号
            "type": "uplink",
            "group_id": group_id,
            "speed": service_ports['uplink']['speed']
        }
        new_connections.append(connection)
    
    # 创建下行端口连接
    # 处理端口范围表示法 (如 "GE1/0/1-48")
    for port_range in service_ports['downlink']['ports']:
        if '-' in port_range:
            base_port = port_range.rsplit('/', 1)[0]
            start, end = map(int, port_range.rsplit('/', 1)[1].split('-'))
            for i in range(start, end + 1):
                port = f"{base_port}/{i}"
                connection = {
                    "id": str(uuid.uuid4()),
                    "source_device": device1_id,
                    "target_device": device2_id,
                    "source_port": port,
                    "target_port": port,  # 使用相同端口号
                    "type": "downlink",
                    "group_id": group_id,
                    "speed": service_ports['downlink']['speed']
                }
                new_connections.append(connection)
        else:
            connection = {
                "id": str(uuid.uuid4()),
                "source_device": device1_id,
                "target_device": device2_id,
                "source_port": port_range,
                "target_port": port_range,  # 使用相同端口号
                "type": "downlink",
                "group_id": group_id,
                "speed": service_ports['downlink']['speed']
            }
            new_connections.append(connection)
    
    return new_connections

@app.route('/create_device_group', methods=['POST'])
def create_device_group():
    try:
        data = request.json
        if not data:
            return jsonify({"error": "No data provided"}), 400

        # 验证必填字段
        if not all(k in data for k in ['name', 'template_id', 'primary_device_id', 'secondary_device_id']):
            return jsonify({"error": "Missing required fields"}), 400

        # 加载数据
        device_groups_data = load_data(DEVICE_GROUPS_FILE)
        device_instances_data = load_data(DEVICE_INSTANCES_FILE)
        templates_data = load_data(GROUP_TEMPLATES_FILE)
        connections_data = load_data(CONNECTIONS_FILE)

        # 验证模板存在
        template = next((t for t in templates_data['templates'] if t['id'] == data['template_id']), None)
        if not template:
            return jsonify({"error": "Template not found"}), 404

        # 验证设备存在
        primary_device = next((d for d in device_instances_data['device_instances'] 
                             if d['id'] == data['primary_device_id']), None)
        secondary_device = next((d for d in device_instances_data['device_instances'] 
                               if d['id'] == data['secondary_device_id']), None)

        if not primary_device or not secondary_device:
            return jsonify({"error": "One or both devices not found"}), 404

        # 验证设备类型匹配
        if primary_device['device_type'] != template['device_type'] or \
           secondary_device['device_type'] != template['device_type']:
            return jsonify({"error": "Device type does not match template"}), 400

        # 验证设备未在其他组中
        for group in device_groups_data['groups']:
            if data['primary_device_id'] in [group['primary_device_id'], group['secondary_device_id']] or \
               data['secondary_device_id'] in [group['primary_device_id'], group['secondary_device_id']]:
                return jsonify({"error": "One or both devices are already in a group"}), 400

        # 创建新组
        new_group = {
            "id": str(uuid.uuid4()),
            "name": data['name'],
            "template_id": data['template_id'],
            "primary_device_id": data['primary_device_id'],
            "secondary_device_id": data['secondary_device_id'],
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat()
        }

        # 添加新组
        device_groups_data['groups'].append(new_group)
        device_groups_data['timestamp'] = datetime.now().isoformat()

        # 更新全局变量
        global device_groups, connections
        device_groups = device_groups_data['groups']

        # 保存设备组数据
        save_data(DEVICE_GROUPS_FILE, device_groups_data['groups'])

        # 自动创建连接
        new_connections = []
        group_id = new_group['id']
        
        # 创建peer-link连接
        for port in template['peer_link_ports']:
            connection = {
                "id": str(uuid.uuid4()),
                "source_device": data['primary_device_id'],
                "target_device": data['secondary_device_id'],
                "source_port": port,
                "target_port": port,
                "type": "peer-link",
                "group_id": group_id,
                "speed": "100G",  # 默认速度，可以根据需要调整
                "created_at": datetime.now().isoformat()
            }
            new_connections.append(connection)
        
        # 创建keepalive连接
        for port in template['keepalive_ports']:
            connection = {
                "id": str(uuid.uuid4()),
                "source_device": data['primary_device_id'],
                "target_device": data['secondary_device_id'],
                "source_port": port,
                "target_port": port,
                "type": "keepalive",
                "group_id": group_id,
                "speed": "25G",  # 默认速度，可以根据需要调整
                "created_at": datetime.now().isoformat()
            }
            new_connections.append(connection)
        
        # 将新连接添加到连接列表
        connections.extend(new_connections)
        save_data(CONNECTIONS_FILE, connections)

        return jsonify({
            "message": "Device group created successfully",
            "group": new_group,
            "connections_created": len(new_connections)
        }), 200

    except Exception as e:
        print(f"Error creating device group: {e}")
        return jsonify({"error": "Internal server error"}), 500

@app.route('/add_device_to_group/<group_id>', methods=['POST'])
def add_device_to_group(group_id):
    """添加设备到组"""
    try:
        data = request.get_json()
        
        if not data or not data.get('device_id'):
            return jsonify({
                'message': '缺少设备ID'
            }), 400
            
        # 查找组和设备
        group = next((g for g in device_groups if g['id'] == group_id), None)
        device = next((d for d in device_instances if d['id'] == data['device_id']), None)
        
        if not group:
            return jsonify({
                'message': '指定的设备组不存在'
            }), 404
            
        if not device:
            return jsonify({
                'message': '指定的设备不存在'
            }), 404
            
        # 检查设备是否已在其他组中
        for g in device_groups:
            if data['device_id'] in [d['id'] for d in g['devices']]:
                return jsonify({
                    'message': '设备已在其他组中'
                }), 400
                
        # 添加设备到组
        group['devices'].append({
            'id': device['id'],
            'role': data.get('role', 'member')  # 可以是 master 或 member
        })
        
        save_data(DEVICE_GROUPS_FILE, device_groups)
        
        return jsonify({
            'message': '设备添加成功',
            'group': group
        }), 200
        
    except Exception as e:
        print(f"Error adding device to group: {str(e)}")
        return jsonify({'message': '添加设备到组时发生错误'}), 500

@app.route('/remove_device_from_group/<group_id>/<device_id>', methods=['DELETE'])
def remove_device_from_group(group_id, device_id):
    """从组中移除设备"""
    try:
        # 查找组
        group = next((g for g in device_groups if g['id'] == group_id), None)
        
        if not group:
            return jsonify({
                'message': '指定的设备组不存在'
            }), 404
            
        # 移除设备
        group['devices'] = [d for d in group['devices'] if d['id'] != device_id]
        
        # 移除相关的连接
        group['peer_links'] = [link for link in group['peer_links'] 
                             if device_id not in [link['device1_id'], link['device2_id']]]
        group['heartbeat_links'] = [link for link in group['heartbeat_links']
                                  if device_id not in [link['device1_id'], link['device2_id']]]
        
        save_data(DEVICE_GROUPS_FILE, device_groups)
        
        return jsonify({
            'message': '设备移除成功',
            'group': group
        }), 200
        
    except Exception as e:
        print(f"Error removing device from group: {str(e)}")
        return jsonify({'message': '从组中移除设备时发生错误'}), 500

@app.route('/configure_group_links/<group_id>', methods=['POST'])
def configure_group_links(group_id):
    """配置组内设备的互连关系"""
    try:
        data = request.get_json()
        
        # 查找组
        group = next((g for g in device_groups if g['id'] == group_id), None)
        
        if not group:
            return jsonify({
                'message': '指定的设备组不存在'
            }), 404
            
        # 验证和更新peer-link配置
        if 'peer_links' in data:
            for link in data['peer_links']:
                if not all([link.get('device1_id'), link.get('device2_id'),
                          link.get('device1_port'), link.get('device2_port')]):
                    return jsonify({
                        'message': 'peer-link配置信息不完整'
                    }), 400
            group['peer_links'] = data['peer_links']
            
        # 验证和更新heartbeat-link配置
        if 'heartbeat_links' in data:
            for link in data['heartbeat_links']:
                if not all([link.get('device1_id'), link.get('device2_id'),
                          link.get('device1_port'), link.get('device2_port')]):
                    return jsonify({
                        'message': 'heartbeat-link配置信息不完整'
                    }), 400
            group['heartbeat_links'] = data['heartbeat_links']
            
        save_data(DEVICE_GROUPS_FILE, device_groups)
        
        return jsonify({
            'message': '组内连接配置成功',
            'group': group
        }), 200
        
    except Exception as e:
        print(f"Error configuring group links: {str(e)}")
        return jsonify({'message': '配置组内连接时发生错误'}), 500

@app.route('/batch_add_connections/<group_id>', methods=['POST'])
def batch_add_connections(group_id):
    """为组内所有设备批量添加连接"""
    try:
        data = request.get_json()
        
        if not data or not all([data.get('target_device'), data.get('target_ports'),
                              data.get('source_ports')]):
            return jsonify({
                'message': '缺少必要的连接信息'
            }), 400
            
        # 查找组
        group = next((g for g in device_groups if g['id'] == group_id), None)
        if not group:
            return jsonify({
                'message': '指定的设备组不存在'
            }), 404
            
        # 验证目标设备
        target_device = next((d for d in device_instances if d['id'] == data['target_device']), None)
        if not target_device:
            return jsonify({
                'message': '目标设备不存在'
            }), 404
            
        # 验证端口数量匹配
        if len(data['target_ports']) != len(data['source_ports']):
            return jsonify({
                'message': '源端口和目标端口数量不匹配'
            }), 400
            
        # 为组内每个设备创建连接
        new_connections = []
        for device in group['devices']:
            source_device = next((d for d in device_instances if d['id'] == device['id']), None)
            if not source_device:
                continue
                
            for source_port, target_port in zip(data['source_ports'], data['target_ports']):
                # 检查端口是否已被占用
                if any(conn['source_device'] == source_device['id'] and conn['source_port'] == source_port
                      or conn['target_device'] == source_device['id'] and conn['target_port'] == source_port
                      for conn in connections):
                    continue
                    
                if any(conn['source_device'] == target_device['id'] and conn['source_port'] == target_port
                      or conn['target_device'] == target_device['id'] and conn['target_port'] == target_port
                      for conn in connections):
                    continue
                    
                # 创建新连接
                new_connection = {
                    'id': str(uuid.uuid4()),
                    'source_device': source_device['id'],
                    'target_device': target_device['id'],
                    'source_port': source_port,
                    'target_port': target_port,
                    'group_id': group_id  # 记录这是组连接
                }
                new_connections.append(new_connection)
                
        # 添加新连接
        connections.extend(new_connections)
        save_data(CONNECTIONS_FILE, connections)
        
        return jsonify({
            'message': '批量连接创建成功',
            'connections': new_connections
        }), 201
        
    except Exception as e:
        print(f"Error creating batch connections: {str(e)}")
        return jsonify({'message': '创建批量连接时发生错误'}), 500

def save_device_port_templates():
    with file_locks['device_port_templates']:
        with open(PORT_TEMPLATES_FILE, 'w', encoding='utf-8') as f:
            json.dump(device_port_templates, f, indent=4, ensure_ascii=False)

@app.route('/get_device_port_templates', methods=['GET'])
def get_device_port_templates():
    try:
        templates_data = load_data(PORT_TEMPLATES_FILE)
        return jsonify(templates_data)
    except Exception as e:
        print(f"Error loading port templates: {e}")
        return jsonify({"templates": {}, "version": "1.0", "last_updated": datetime.now().strftime("%Y-%m-%d")})

# 更新设备端口模板
@app.route('/update_device_port_template', methods=['POST'])
def update_device_port_template():
    try:
        data = request.get_json()
        device_type = data.get('device_type')
        template = data.get('template')
        
        if not device_type or not template:
            return jsonify({"error": "Missing required fields"}), 400
            
        templates_data = load_data(PORT_TEMPLATES_FILE)
        templates = templates_data.get("templates", {})
        templates[device_type] = template
        templates_data["templates"] = templates
        templates_data["last_updated"] = datetime.now().strftime("%Y-%m-%d")
        
        save_data(PORT_TEMPLATES_FILE, templates_data)
        return jsonify({"message": "模板保存成功"})
    except Exception as e:
        print(f"Error updating port template: {e}")
        return jsonify({"error": "Failed to update template"}), 500

# 删除设备端口模板
@app.route('/delete_device_port_template/<device_type>', methods=['DELETE'])
def delete_device_port_template(device_type):
    try:
        templates_data = load_data(PORT_TEMPLATES_FILE)
        templates = templates_data.get("templates", {})
        
        if device_type in templates:
            del templates[device_type]
            templates_data["templates"] = templates
            templates_data["last_updated"] = datetime.now().strftime("%Y-%m-%d")
            save_data(PORT_TEMPLATES_FILE, templates_data)
            return jsonify({"message": "模板删除成功"})
        else:
            return jsonify({"error": "Template not found"}), 404
    except Exception as e:
        print(f"Error deleting port template: {e}")
        return jsonify({"error": "Failed to delete template"}), 500

# 添加辅助函数来检查端口是否在同类型组中被使用
def is_port_used_in_same_group_type(device_id, port_name, group_type):
    """
    检查端口是否在同类型的组中被使用
    :param device_id: 设备ID
    :param port_name: 端口名称
    :param group_type: 组类型 (stack/mlag)
    :return: bool
    """
    # 查找该设备所有相关的连接
    device_connections = [conn for conn in connections 
        if (conn['source_device'] == device_id and conn['source_port'] == port_name) or 
           (conn['target_device'] == device_id and conn['target_port'] == port_name)]
    
    for conn in device_connections:
        # 如果连接属于某个组
        if 'group_id' in conn:
            # 查找该组的类型
            group = next((g for g in device_groups if g['id'] == conn['group_id']), None)
            if group and group['type'] == group_type:
                return True
    return False

# Group Template Management Routes
@app.route('/get_group_templates')
def get_group_templates():
    try:
        data = load_data(GROUP_TEMPLATES_FILE)
        return jsonify(data)
    except Exception as e:
        print(f"Error getting group templates: {e}")
        return jsonify({"message": "获取组模板失败", "error": str(e)}), 500

@app.route('/create_group_template', methods=['POST'])
def create_group_template():
    try:
        data = load_data(GROUP_TEMPLATES_FILE)
        template = request.json
        
        # Generate a unique ID for the new template
        template_id = str(uuid.uuid4())
        template['id'] = template_id
        
        # Add timestamp
        template['created_at'] = datetime.now().isoformat()
        template['updated_at'] = template['created_at']
        
        # Add the new template to the list
        data['templates'].append(template)
        
        # Update timestamp
        data['timestamp'] = datetime.now().isoformat()
        
        # Save the updated data
        save_data(GROUP_TEMPLATES_FILE, data)
        
        return jsonify({"message": "组模板创建成功", "template_id": template_id})
    except Exception as e:
        print(f"Error creating group template: {e}")
        return jsonify({"message": "创建组模板失败", "error": str(e)}), 500

@app.route('/update_group_template/<template_id>', methods=['PUT'])
def update_group_template(template_id):
    try:
        data = load_data(GROUP_TEMPLATES_FILE)
        template = request.json
        
        # Find and update the template
        for i, t in enumerate(data['templates']):
            if t['id'] == template_id:
                template['id'] = template_id
                template['created_at'] = t['created_at']
                template['updated_at'] = datetime.now().isoformat()
                data['templates'][i] = template
                break
        else:
            return jsonify({"message": "未找到指定的组模板"}), 404
        
        # Update timestamp
        data['timestamp'] = datetime.now().isoformat()
        
        # Save the updated data
        save_data(GROUP_TEMPLATES_FILE, data)
        
        return jsonify({"message": "组模板更新成功"})
    except Exception as e:
        print(f"Error updating group template: {e}")
        return jsonify({"message": "更新组模板失败", "error": str(e)}), 500

@app.route('/delete_group_template/<template_id>', methods=['DELETE'])
def delete_group_template(template_id):
    try:
        data = load_data(GROUP_TEMPLATES_FILE)
        
        # Find and remove the template
        data['templates'] = [t for t in data['templates'] if t['id'] != template_id]
        
        # Update timestamp
        data['timestamp'] = datetime.now().isoformat()
        
        # Save the updated data
        save_data(GROUP_TEMPLATES_FILE, data)
        
        return jsonify({"message": "组模板删除成功"})
    except Exception as e:
        print(f"Error deleting group template: {e}")
        return jsonify({"message": "删除组模板失败", "error": str(e)}), 500

# Device Group Management Routes
@app.route('/group_templates')
def group_templates_page():
    return render_template('group_template_management.html')

@app.route('/device_groups')
def device_groups_page():
    return render_template('device_group_management.html')

def save_connections(connections_data):
    """保存连接数据到文件"""
    try:
        with file_lock:
            # 确保数据目录存在
            ensure_data_dir()
            
            # 读取现有数据
            data = load_data(CONNECTIONS_FILE)
            
            # 更新连接数据
            data['connections'] = connections_data
            data['timestamp'] = datetime.now().isoformat()
            
            # 保存到文件
            with open(CONNECTIONS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            
            # 创建备份
            backup_file = os.path.join(DATA_DIR, f'connections.{int(time.time())}.json')
            with open(backup_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            
            return True
    except Exception as e:
        print(f"Error saving connections: {e}")
        return False

def save_device_groups(groups_data):
    """保存设备组数据到文件"""
    try:
        with file_lock:
            # 确保数据目录存在
            ensure_data_dir()
            
            # 读取现有数据
            data = load_data(DEVICE_GROUPS_FILE)
            
            # 更新设备组数据
            data['groups'] = groups_data
            data['timestamp'] = datetime.now().isoformat()
            
            # 保存到文件
            with open(DEVICE_GROUPS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            
            # 创建备份
            backup_file = os.path.join(DATA_DIR, f'device_groups.{int(time.time())}.json')
            with open(backup_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=4)
            
            return True
    except Exception as e:
        print(f"Error saving device groups: {e}")
        return False

if __name__ == '__main__':
    app.run(host='172.31.60.204', port=58000, debug=True)
