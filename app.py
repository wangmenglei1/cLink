from flask import Flask, render_template, request, jsonify, redirect, url_for, flash, send_file, g
import os
import json
import time
import uuid
from datetime import datetime
import threading
import sys
import shutil

# Import pandas and io for Excel export
import pandas as pd
import io
from openpyxl import Workbook
from openpyxl.styles import Border, Side, PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

# --- Interface Configuration ---
# A structured configuration for defining interfaces based on rate and media type.
INTERFACE_CONFIG = {
    "1G": {
        "media_types": {
            "电口": {"prefix": "GE", "type": "electric"},
            "光口": {"prefix": "GE", "type": "optic"}
        },
        "default_media": "电口"
    },
    "10G": {
        "media_types": {
            "电口": {"prefix": "XGE", "type": "electric"},
            "光口": {"prefix": "XGE", "type": "optic"}
        },
        "default_media": "光口"
    },
    "25G": {
        "media_types": {
            "光口": {"prefix": "TGE", "type": "optic"}
        },
        "default_media": "光口"
    },
    "40G": {
        "media_types": {
            "光口": {"prefix": "FGE", "type": "optic"}
        },
        "default_media": "光口"
    },
    "100G": {
        "media_types": {
            "光口": {"prefix": "HGE", "type": "optic"}
        },
        "default_media": "光口"
    },
    "管理口": {
        "media_types": {
            "管理口": {"prefix": "Mgmt", "type": "management"}
        },
        "default_media": "管理口"
    },
    "堆叠口": {
        "media_types": {
            "堆叠口": {"prefix": "Stack", "type": "stack"}
        },
        "default_media": "堆叠口"
    }
}

# 文件锁对象
file_locks = {
    'device_instances': threading.Lock(),
    'racks': threading.Lock(),
    'rooms': threading.Lock(),
    'device_types': threading.Lock(),
    'connections': threading.Lock()
}

app = Flask(__name__)

# --- Data Storage --- #
DATA_DIR = 'data'
DEVICE_TYPES_FILE = os.path.join(DATA_DIR, 'device_types.json')
DEVICE_INSTANCES_FILE = os.path.join(DATA_DIR, 'device_instances.json')
CONNECTIONS_FILE = os.path.join(DATA_DIR, 'connections.json')
ROOMS_FILE = os.path.join(DATA_DIR, 'rooms.json') # New file for rooms
RACKS_FILE = os.path.join(DATA_DIR, 'racks.json') # New file for racks

# In-memory storage (will be loaded from file)
device_types = []
device_instances = []
connections = []
rooms = [] # New list for rooms
racks = [] # New list for racks

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
        return {'device_types': []}
    elif 'connections' in filepath:
        return {'connections': []}
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
        else:
            lock = threading.Lock()
        
        with lock:
            # 使用临时文件确保原子性写入
            temp_file = filepath + '.tmp'
            
            # 包装数据为新格式，添加版本信息和时间戳
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
                data_wrapper = {
                    "version": "2.0",
                    "timestamp": datetime.now().isoformat(),
                    "connections": data
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
    global device_types, device_instances, connections, rooms, racks
    
    try:
        # Load all data files with corruption recovery
        device_types_data = load_data(DEVICE_TYPES_FILE)
        device_instances_data = load_data(DEVICE_INSTANCES_FILE)
        connections_data = load_data(CONNECTIONS_FILE)
        rooms_data = load_data(ROOMS_FILE)
        racks_data = load_data(RACKS_FILE)
        
        # Extract data from new format (with fallback for old format)
        # 先检查是否为列表（旧格式），再尝试新格式
        device_types = device_types_data if isinstance(device_types_data, list) else device_types_data.get('device_types', [])
        device_instances = device_instances_data if isinstance(device_instances_data, list) else device_instances_data.get('device_instances', [])
        connections = connections_data if isinstance(connections_data, list) else connections_data.get('connections', [])
        rooms = rooms_data if isinstance(rooms_data, list) else rooms_data.get('rooms', [])
        racks = racks_data if isinstance(racks_data, list) else racks_data.get('racks', [])
        
        print(f"Data loaded - Rooms: {len(rooms)}, Racks: {len(racks)}, Devices: {len(device_instances)}, Types: {len(device_types)}, Connections: {len(connections)}")
        
    except Exception as e:
        print(f"Error loading data: {e}")
        # Initialize with empty data if loading fails
        device_types = []
        device_instances = []
        connections = []
        rooms = []
        racks = []

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
    
    # 对每个排内的机柜按编号排序
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
        
        # Create a set of used type names for efficient lookup
        used_types = {instance['device_type'] for instance in device_instances}
        
        # Augment device types with usage info
        augmented_device_types = []
        for type_data in device_types:
            # Make a copy to avoid modifying the global list
            type_info = type_data.copy()
            type_info['is_in_use'] = type_info['name'] in used_types
            augmented_device_types.append(type_info)

        return jsonify({'device_types': augmented_device_types}), 200
    except Exception as e:
        print(f"Error fetching device types: {e}")
        return jsonify({'message': 'Failed to fetch device types'}), 500

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
            media_type = interface_group.get('media_type')
            naming_prefix = interface_group.get('naming_prefix', '1/0/')
            start = interface_group.get('start')
            end = interface_group.get('end')

            rate_config = INTERFACE_CONFIG.get(rate)
            if not rate_config:
                return jsonify({'message': f'无效的速率: {rate}'}), 400
            
            media_config = rate_config['media_types'].get(media_type)
            if not media_config:
                return jsonify({'message': f'速率 "{rate}" 不支持介质类型 "{media_type}"'}), 400

            if not all([isinstance(start, int), isinstance(end, int)]):
                return jsonify({'message': '接口组缺少参数或类型错误。'}), 400
            if start > end:
                return jsonify({'message': f'接口组 "{rate} - {media_type}" 的起始编号不能大于结束编号。'}), 400

            base_prefix = media_config['prefix']
            if_type = media_config['type']

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
            return jsonify({'message': '要更新的设备型号未找到。'}), 404

        # Prevent editing if the type is in use
        if any(instance['device_type'] == original_type_name for instance in device_instances):
            return jsonify({'message': f'设备型号 "{original_type_name}" 正在使用中，无法修改。'}), 400

        data = request.get_json()
        new_name = data.get('deviceModelName', '').strip()
        new_height = data.get('height_u')
        new_definitions = data.get('interfaces', [])

        # Validate new data
        if not new_name or not isinstance(new_height, int) or new_height <= 0:
            return jsonify({'message': '设备型号名称和有效高度为必填项。'}), 400
        
        # Check for name conflict if the name is being changed
        if new_name != original_type_name and any(t['name'] == new_name for t in device_types):
            return jsonify({'message': f'设备型号名称 "{new_name}" 已存在。'}), 409

        # Generate new full interface list from structured definitions
        new_full_interfaces_list = []
        for interface_group in new_definitions:
            rate = interface_group.get('rate')
            media_type = interface_group.get('media_type')
            naming_prefix = interface_group.get('naming_prefix', '1/0/')
            start = interface_group.get('start')
            end = interface_group.get('end')

            rate_config = INTERFACE_CONFIG.get(rate)
            if not rate_config:
                return jsonify({'message': f'无效的速率: {rate}'}), 400

            media_config = rate_config['media_types'].get(media_type)
            if not media_config:
                return jsonify({'message': f'速率 "{rate}" 不支持介质类型 "{media_type}"'}), 400

            if not all([isinstance(start, int), isinstance(end, int)]):
                return jsonify({'message': '接口组数据无效。'}), 400
            
            base_prefix = media_config['prefix']
            if_type = media_config['type']

            for i in range(start, end + 1):
                new_full_interfaces_list.append({'name': f"{base_prefix}{naming_prefix}{i}", 'type': if_type, 'connected_to': None})

        # Update the device type
        target_type['name'] = new_name
        target_type['height_u'] = new_height
        target_type['interfaces'] = new_full_interfaces_list
        target_type['definitions'] = new_definitions

        save_data(DEVICE_TYPES_FILE, device_types)
        return jsonify({'message': '设备型号更新成功。'}), 200

    except Exception as e:
        print(f"Error updating device type: {str(e)}")
        return jsonify({'message': '更新设备型号时发生内部错误。'}), 500

@app.route('/delete_device_type/<string:device_type_name>', methods=['DELETE'])
def delete_device_type(device_type_name):
    """Deletes a device type from the system."""
    try:
        load_all_data_before_request()

        # Check if the device type is in use before deleting
        if any(instance['device_type'] == device_type_name for instance in device_instances):
            return jsonify({'message': f'设备型号 "{device_type_name}" 正在被一个或多个设备实例使用，无法删除。'}), 400

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
        return jsonify({'message': '设备型号、命名前缀、起始和结束编号均为必填项。'}), 400
    
    if start_number > end_number:
        return jsonify({'message': '起始编号不能大于结束编号。'}), 400

    if end_number - start_number + 1 > 200: # Limit batch size to prevent abuse/overload
        return jsonify({'message': '单次批量添加的设备数量不能超过200台。'}), 400
    
    load_all_data_before_request()

    device_type_definition = next((dt for dt in device_types if dt['name'] == device_type_name), None)
    if not device_type_definition:
        return jsonify({'message': f'设备型号 "{device_type_name}" 未找到。'}), 404

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
        return jsonify({'message': f'未能添加任何新设备。以下名称均已存在: {", ".join(skipped_names)}'}), 409

    device_instances.extend(new_instances)
    save_data(DEVICE_INSTANCES_FILE, device_instances)

    message = f'成功添加 {len(new_instances)} 台设备。'
    if skipped_names:
        message += f' 跳过 {len(skipped_names)} 台已存在的设备: {", ".join(skipped_names)}'
    
    return jsonify({'message': message, 'added_count': len(new_instances), 'skipped_count': len(skipped_names)}), 201

@app.route('/add_connection', methods=['POST'])
def add_connection():
    try:
        data = request.get_json()
        if not data:
            return jsonify({'error': 'Invalid JSON', 'message': 'No JSON data received'}), 400

        source_id = data.get('source')
        target_id = data.get('target')
        source_port_name = data.get('source_port')
        target_port_name = data.get('target_port')

        if not all([source_id, target_id, source_port_name, target_port_name]):
             return jsonify({'error': 'Missing data', 'message': 'Missing source, target, source_port, or target_port for connection'}), 400

        source_instance = find_instance_by_id(source_id)
        target_instance = find_instance_by_id(target_id)

        if not source_instance or not target_instance:
            return jsonify({'error': 'Invalid device instance ID', 'message': 'Source or target device instance not found'}), 400

        source_interface = next((iface for iface in source_instance['interfaces'] if iface['name'] == source_port_name), None)
        target_interface = next((iface for iface in target_instance['interfaces'] if iface['name'] == target_port_name), None)

        if not source_interface or not target_interface:
             return jsonify({'error': 'Invalid port name', 'message': 'Source or target port not found on the device instance'}), 400

        if source_interface['status'] != 'available' or target_interface['status'] != 'available':
             return jsonify({'error': 'Port already in use', 'message': 'One or both selected ports are already connected'}), 400

        # Update interface status to connected
        update_interface_status(source_id, source_port_name, 'connected')
        update_interface_status(target_id, target_port_name, 'connected')

        new_connection = {
            'source': source_id,
            'target': target_id,
            'source_port': source_port_name,
            'target_port': target_port_name,
        }
        connections.append(new_connection)

        # Save connections and updated instances
        save_data(CONNECTIONS_FILE, connections)
        save_data(DEVICE_INSTANCES_FILE, device_instances)

        # print("Received new connection:", new_connection) # Avoid excessive printing
        # print("Current connections:", connections) # Avoid excessive printing

        # Return updated instance data so frontend can refresh interface lists
        return jsonify({'message': 'Connection added successfully', 'connection': new_connection, 'updated_instances': [source_instance, target_instance]}), 201

    except Exception as e:
        print(f'Error adding connection: {e}')
        return jsonify({'message': 'Internal server error during connection addition'}), 500

# Route to delete a device instance
@app.route('/delete_device_instance/<string:instance_id>', methods=['DELETE'])
def delete_device_instance(instance_id):
    try:
        load_all_data_before_request()

        # Check if the device is still mounted
        instance = find_instance_by_id(instance_id)
        if instance and instance.get('rack_id'):
            return jsonify({'message': '设备必须先下架才能删除。'}), 400

        success, updated_instances_data = perform_delete_device_instance(instance_id)

        if success:
            return jsonify({'message': f'Device instance {instance_id} and associated connections deleted successfully', 'deleted_instance_id': instance_id, 'updated_instances': updated_instances_data}), 200
        else:
             # This case should ideally not be reached if instance_to_delete was checked in perform_delete_device_instance
             return jsonify({'error': 'Not Found', 'message': f'Device instance with ID {instance_id} not found'}), 404

    except Exception as e:
        print(f'Error deleting device instance: {e}')
        return jsonify({'message': 'Internal server error during device instance deletion'}), 500

# Route to delete a connection
@app.route('/delete_connection', methods=['POST']) # Using POST as DELETE typically doesn't have body
def delete_connection():
    global connections, device_instances
    data = request.get_json()
    if not data:
        return jsonify({'error': 'Invalid JSON', 'message': 'No JSON data received'}), 400

    source_id = data.get('source')
    target_id = data.get('target')
    source_port_name = data.get('source_port')
    target_port_name = data.get('target_port')

    if not all([source_id, target_id, source_port_name, target_port_name]):
         return jsonify({'error': 'Missing data', 'message': 'Missing source, target, source_port, or target_port for connection deletion'}), 400

    # Find the connection to be deleted
    conn_to_delete = find_connection(source_id, target_id, source_port_name, target_port_name)
    
    if not conn_to_delete:
         return jsonify({'error': 'Not Found', 'message': 'Connection not found'}), 404

    # Remove the connection from the global list
    initial_connection_count = len(connections)
    connections = [conn for conn in connections if conn != conn_to_delete]
    
    # Update interface status on connected instances
    updated_instances_data = []

    # Free up ports on source and target instances
    if update_interface_status(source_id, source_port_name, 'available'):
         updated_source_instance = find_instance_by_id(source_id)
         if updated_source_instance: # Ensure instance exists before adding
              updated_instances_data.append(updated_source_instance)

    if update_interface_status(target_id, target_port_name, 'available'):
         updated_target_instance = find_instance_by_id(target_id)
         # Prevent adding the same instance twice if source and target are the same
         if updated_target_instance and updated_target_instance not in updated_instances_data: 
              updated_instances_data.append(updated_target_instance)


    # Save updated data to files
    save_data(CONNECTIONS_FILE, connections)
    save_data(DEVICE_INSTANCES_FILE, device_instances) # Save instances because interface status changed

    print(f"Deleted connection: {source_id}:{source_port_name} - {target_id}:{target_port_name}. Removed {initial_connection_count - len(connections)} connection(s). Affected instances updated: {[inst['id'] for inst in updated_instances_data]}.")

     # Return updated instance data so frontend can refresh interface lists
    return jsonify({'message': 'Connection deleted successfully', 'updated_instances': updated_instances_data}), 200

@app.route('/get_device_instances', methods=['GET'])
def get_device_instances():
    return jsonify({'device_instances': device_instances}), 200

@app.route('/get_connections', methods=['GET'])
def get_connections():
    return jsonify({'connections': connections}), 200

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
                return jsonify({'message': '机柜高度必须是1-100之间的正整数'}), 400
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

@app.route('/update_device_instance_rack_position/<string:instance_id>', methods=['PUT'])
def update_device_instance_rack_position(instance_id):
    try:
        # 不要重新加载数据，直接使用内存中的数据
        # load_all_data_before_request()
        
        data = request.get_json()

        if not data:
            return jsonify({'message': 'Invalid JSON data received'}), 400

        rack_id = data.get('rack_id')
        rack_u = data.get('rack_u')

        if rack_id is None or rack_u is None:
             return jsonify({'message': 'Missing rack_id or rack_u in request data.'}), 400

        # Find the device instance
        instance = find_instance_by_id(instance_id)
        if not instance:
             return jsonify({'message': f'Device instance with ID {instance_id} not found'}), 404

        # Find the target rack
        target_rack = next((r for r in racks if r['id'] == rack_id), None)
        if not target_rack:
             return jsonify({'message': f'Rack with ID {rack_id} not found'}), 404

        # Basic validation for rack_u (must be positive integer)
        if not isinstance(rack_u, int) or rack_u <= 0:
             return jsonify({'message': 'Invalid rack_u. Must be a positive integer.'}), 400
             
        # TODO: Add more advanced validation here:
        # - Check if device height + rack_u is within rack height of target_rack
        # - Check for U position conflicts with other devices already in target_rack
        # - Check if the device is already in a rack and handle removing it from the old rack's device list

        # If the instance was previously in a rack, remove it from that rack's devices list
        if 'rack_id' in instance and instance['rack_id'] is not None and instance['rack_id'] != rack_id:
             old_rack = next((r for r in racks if r['id'] == instance['rack_id']), None)
             if old_rack and 'devices' in old_rack:
                 old_rack['devices'] = [dev_id for dev_id in old_rack['devices'] if dev_id != instance_id]
                 # Note: We will save racks data once after all updates

        # Update the instance's rack information
        instance['rack_id'] = rack_id
        instance['rack_u'] = rack_u

        # Add the instance ID to the new target rack's devices list
        if 'devices' not in target_rack:
            target_rack['devices'] = [] # Initialize if it doesn't exist
        if instance_id not in target_rack['devices']:
             target_rack['devices'].append(instance_id)

        # Save updated data immediately
        save_data(DEVICE_INSTANCES_FILE, device_instances)
        save_data(RACKS_FILE, racks) # Save racks data after potentially modifying old and new rack lists

        print(f'[API] update_rack_position: Updated rack position for instance {instance_id} to Rack {rack_id} at U {rack_u}')
        
        # 验证数据是否正确保存
        saved_device_instances_data = load_data(DEVICE_INSTANCES_FILE)
        saved_instance = next((inst for inst in saved_device_instances_data if inst['id'] == instance_id), None)
        if saved_instance and saved_instance.get('rack_id') == rack_id:
            print(f'✓ Data verification successful: instance {instance_id} correctly saved with rack_id {rack_id}')
        else:
            print(f'✗ Data verification failed: instance {instance_id} not correctly saved!')
            
        return jsonify({'message': 'Device instance rack position updated successfully', 'instance_id': instance_id, 'rack_id': rack_id, 'rack_u': rack_u}), 200

    except Exception as e:
        print(f'Error updating device instance rack position: {e}')
        return jsonify({'message': 'Internal server error during rack position update'}), 500

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
                
            # 计算本排的最大U数
            max_u = max(rack['height_u'] for rack in racks_in_row)
            n_racks = len(racks_in_row)
            
            # 每个机柜4列：左U数、设备名、右U数、间隔
            total_cols = 1 + n_racks * 4
            
            # 设置列宽
            ws.column_dimensions[get_column_letter(current_col)].width = 6  # 排名列
            for i, rack in enumerate(racks_in_row):
                col_offset = current_col + 1 + i * 4
                ws.column_dimensions[get_column_letter(col_offset)].width = 4  # 左U数列
                ws.column_dimensions[get_column_letter(col_offset + 1)].width = 20  # 设备列
                ws.column_dimensions[get_column_letter(col_offset + 2)].width = 4  # 右U数列
                ws.column_dimensions[get_column_letter(col_offset + 3)].width = 3  # 间隔列
            
            # 左侧排名合并单元格
            ws.merge_cells(start_row=current_row, start_column=current_col, 
                          end_row=current_row + max_u, end_column=current_col)
            cell = ws.cell(row=current_row, column=current_col)
            cell.value = f'{row}排'
            cell.border = normal_border_style
            cell.fill = row_title_fill
            cell.alignment = center_alignment
            
            # 机柜标题行
            for i, rack in enumerate(racks_in_row):
                col_offset = current_col + 1 + i * 4
                
                # 左U列标题
                u_cell = ws.cell(row=current_row, column=col_offset)
                u_cell.value = 'U'
                u_cell.border = normal_border_style
                u_cell.alignment = center_alignment
                
                # 机柜名标题
                rack_dep = rack.get('department')
                rack_cell = ws.cell(row=current_row, column=col_offset + 1)
                rack_cell.value = rack['name'] if not rack_dep else f"{rack['name']} ({rack_dep})"
                rack_cell.border = normal_border_style
                rack_cell.alignment = center_alignment
                
                # 右U列标题
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
            
            # 生成U数和设备块
            for u_idx in range(max_u):
                # U编号从1到max_u，底部是1
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
                    
                    # 写入设备块
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
                        # 空U位
                        empty_cell = ws.cell(row=row_pos, column=col_offset + 1)
                        empty_cell.value = ''
                        empty_cell.border = normal_border_style
                        empty_cell.alignment = center_alignment
            
            # ✨ 关键功能：为整个表格区域设置外边框（粗边框）
            table_start_row = current_row
            table_end_row = current_row + max_u
            table_start_col = current_col
            table_end_col = current_col + total_cols - 1
            
            # 设置外边框 - 只影响边界，不影响内部格线
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
    try:
        load_all_data_before_request()
        data = request.get_json()
        
        if not data:
            return jsonify({'message': 'Invalid JSON data received'}), 400
            
        rack_id = data.get('rack_id')
        device_updates = data.get('device_updates', [])  # [{'device_id': 'xxx', 'rack_u': 5}, ...]
        
        if not rack_id or not device_updates:
            return jsonify({'message': 'Missing rack_id or device_updates'}), 400
            
        # Find the target rack
        rack = next((r for r in racks if r['id'] == rack_id), None)
        if not rack:
            return jsonify({'message': f'Rack with ID {rack_id} not found'}), 404
            
        # === 第一阶段：完整验证，确保所有操作都能成功 ===
        validation_errors = []
        
        # 验证所有设备存在且未分配
        for update in device_updates:
            device_id = update['device_id']
            rack_u = update['rack_u']
            
            # Check device exists
            device = find_instance_by_id(device_id)
            if not device:
                validation_errors.append(f'Device {device_id} not found')
                continue
                
            # Check device is not already in a rack
            if device.get('rack_id'):
                validation_errors.append(f'Device {device_id} already assigned to rack {device["rack_id"]}')
                continue
                
            # Validate rack_u
            if not isinstance(rack_u, int) or rack_u <= 0:
                validation_errors.append(f'Invalid rack_u {rack_u} for device {device_id}')
                continue
                
            # Get device height and check space
            device_type = next((dt for dt in device_types if dt['name'] == device['device_type']), None)
            if device_type:
                device_height = device_type.get('height_u', 1)
                if rack_u + device_height - 1 > rack['height_u']:
                    validation_errors.append(f'Device {device_id} exceeds rack height at U{rack_u}')
                    
        if validation_errors:
            return jsonify({'message': 'Pre-validation failed - no devices were mounted', 'errors': validation_errors}), 400
            
        # === 检查空间冲突（包括现有设备和批次内冲突） ===
        devices_in_rack = [inst for inst in device_instances if inst.get('rack_id') == rack_id]
        conflict_errors = []
        
        # 构建占用空间映射
        occupied_positions = set()
        
        # 添加现有设备占用的位置
        for existing_device in devices_in_rack:
            if existing_device.get('rack_u'):
                existing_device_type = next((dt for dt in device_types if dt['name'] == existing_device['device_type']), None)
                existing_height = existing_device_type.get('height_u', 1) if existing_device_type else 1
                existing_start = existing_device['rack_u']
                
                for u in range(existing_start, existing_start + existing_height):
                    occupied_positions.add(u)
        
        # 验证批次内无冲突，并检查与现有设备无冲突
        batch_positions = set()
        for update in device_updates:
            device_id = update['device_id']
            rack_u = update['rack_u']
            device = find_instance_by_id(device_id)
            device_type = next((dt for dt in device_types if dt['name'] == device['device_type']), None)
            device_height = device_type.get('height_u', 1) if device_type else 1
            
            # 检查该设备要占用的所有U位
            device_positions = set(range(rack_u, rack_u + device_height))
            
            # 检查与现有设备冲突
            if device_positions & occupied_positions:
                conflict_errors.append(f'Device {device_id} (U{rack_u}-{rack_u + device_height - 1}) conflicts with existing devices')
                continue
                
            # 检查与批次内其他设备冲突
            if device_positions & batch_positions:
                conflict_errors.append(f'Device {device_id} (U{rack_u}-{rack_u + device_height - 1}) conflicts with other devices in this batch')
                continue
                
            # 添加到批次占用位置
            batch_positions.update(device_positions)
                    
        if conflict_errors:
            return jsonify({'message': 'Space conflict detected - no devices were mounted', 'errors': conflict_errors}), 400
            
        # === 第二阶段：所有验证通过，执行原子操作 ===
        print(f'Starting atomic batch mount operation for {len(device_updates)} devices in rack {rack_id}')
        
        # 备份原始状态（用于回滚）
        original_device_states = {}
        for update in device_updates:
            device = find_instance_by_id(update['device_id'])
            original_device_states[update['device_id']] = {
                'rack_id': device.get('rack_id'),
                'rack_u': device.get('rack_u')
            }
        
        original_rack_devices = list(rack.get('devices', []))
        
        try:
            # 更新所有设备
            updated_device_ids = []
            for update in device_updates:
                device_id = update['device_id']
                rack_u = update['rack_u']
                device = find_instance_by_id(device_id)
                
                device['rack_id'] = rack_id
                device['rack_u'] = rack_u
                updated_device_ids.append(device_id)
            
            # 更新机柜设备列表
            if 'devices' not in rack:
                rack['devices'] = []
            rack['devices'].extend(updated_device_ids)
            
            # 原子保存所有数据
            save_data(DEVICE_INSTANCES_FILE, device_instances)
            save_data(RACKS_FILE, racks)
            
            print(f'Batch rack mount completed successfully: {len(updated_device_ids)} devices added to rack {rack_id}')
            return jsonify({
                'message': f'Batch mount successful - all {len(updated_device_ids)} devices mounted',
                'success_count': len(updated_device_ids),
                'rack_id': rack_id,
                'updated_devices': updated_device_ids
            }), 200
            
        except Exception as save_error:
            # 回滚操作
            print(f'Save failed during batch mount, rolling back: {save_error}')
            
            # 恢复设备状态
            for device_id, original_state in original_device_states.items():
                device = find_instance_by_id(device_id)
                device['rack_id'] = original_state['rack_id']
                device['rack_u'] = original_state['rack_u']
            
            # 恢复机柜设备列表
            rack['devices'] = original_rack_devices
            
            raise save_error
        
    except Exception as e:
        print(f'Error in batch rack mount: {e}')
        return jsonify({'message': f'Batch mount failed - no devices were mounted: {str(e)}'}), 500

# --- Multi-Rack Batch Mount API for Sequential Processing ---
@app.route('/multi_rack_batch_mount', methods=['POST'])
def multi_rack_batch_mount():
    """
    多机柜批量上架API - 串行处理避免并发冲突
    请求格式: {
        "mount_requests": [
            {
                "rack_id": "rack_1",
                "device_updates": [{"device_id": "dev1", "rack_u": 1}]
            },
            {
                "rack_id": "rack_2", 
                "device_updates": [{"device_id": "dev2", "rack_u": 1}]
            }
        ]
    }
    """
    try:
        load_all_data_before_request()
        data = request.get_json()
        
        if not data or 'mount_requests' not in data:
            return jsonify({'message': 'Invalid JSON data - mount_requests required'}), 400
            
        mount_requests = data['mount_requests']
        if not isinstance(mount_requests, list) or not mount_requests:
            return jsonify({'message': 'mount_requests must be a non-empty list'}), 400
        
        # 统计信息
        total_racks = len(mount_requests)
        total_devices = sum(len(req.get('device_updates', [])) for req in mount_requests)
        
        print(f'Starting multi-rack batch mount: {total_racks} racks, {total_devices} devices')
        
        # 结果统计
        successful_racks = 0
        failed_racks = 0
        successful_devices = 0
        failed_devices = 0
        errors = []
        
        # 按机柜串行处理，避免并发写入
        for i, mount_request in enumerate(mount_requests, 1):
            rack_id = mount_request.get('rack_id')
            device_updates = mount_request.get('device_updates', [])
            
            print(f'Processing rack {i}/{total_racks}: {rack_id} with {len(device_updates)} devices')
            
            # 为每个机柜创建独立的批量上架请求
            single_rack_data = {
                'rack_id': rack_id,
                'device_updates': device_updates
            }
            
            try:
                # 重新加载数据确保数据最新
                load_all_data_before_request()
                
                # 调用单机柜批量上架逻辑
                rack = next((r for r in racks if r['id'] == rack_id), None)
                if not rack:
                    error_msg = f'Rack {rack_id} not found'
                    errors.append(error_msg)
                    failed_racks += 1
                    failed_devices += len(device_updates)
                    print(f'❌ Rack {i}/{total_racks} failed: {error_msg}')
                    continue
                
                # 执行验证和上架（复用单机柜逻辑）
                result = process_single_rack_mount(rack_id, device_updates)
                
                if result['success']:
                    successful_racks += 1
                    successful_devices += result['device_count']
                    print(f'✅ Rack {i}/{total_racks} successful: {result["device_count"]} devices mounted')
                else:
                    failed_racks += 1
                    failed_devices += len(device_updates)
                    errors.append(f'Rack {rack_id}: {result["error"]}')
                    print(f'❌ Rack {i}/{total_racks} failed: {result["error"]}')
                
                # 短暂延迟确保文件操作完成
                time.sleep(0.1)
                
            except Exception as e:
                error_msg = f'Rack {rack_id}: {str(e)}'
                errors.append(error_msg)
                failed_racks += 1
                failed_devices += len(device_updates)
                print(f'❌ Rack {i}/{total_racks} error: {str(e)}')
        
        # 生成结果报告
        result_message = f'Multi-rack batch mount completed: {successful_racks}/{total_racks} racks successful, {successful_devices}/{total_devices} devices mounted'
        
        response_data = {
            'message': result_message,
            'total_racks': total_racks,
            'successful_racks': successful_racks,
            'failed_racks': failed_racks,
            'total_devices': total_devices,
            'successful_devices': successful_devices,
            'failed_devices': failed_devices,
            'success_rate': f'{(successful_devices/total_devices*100):.1f}%' if total_devices > 0 else '0%',
            'errors': errors
        }
        
        print(f'Multi-rack batch mount result: {response_data}')
        
        # 根据成功率决定HTTP状态码
        if successful_racks == total_racks:
            return jsonify(response_data), 200  # 全部成功
        elif successful_racks > 0:
            return jsonify(response_data), 206  # 部分成功
        else:
            return jsonify(response_data), 400  # 全部失败
            
    except Exception as e:
        print(f'Error in multi-rack batch mount: {e}')
        return jsonify({'message': f'Multi-rack batch mount failed: {str(e)}'}), 500

def process_single_rack_mount(rack_id, device_updates):
    """
    处理单个机柜的批量上架逻辑（从batch_rack_mount提取）
    返回: {'success': bool, 'device_count': int, 'error': str}
    """
    try:
        # 找到目标机柜
        rack = next((r for r in racks if r['id'] == rack_id), None)
        if not rack:
            return {'success': False, 'device_count': 0, 'error': f'Rack {rack_id} not found'}
        
        if not device_updates:
            return {'success': True, 'device_count': 0, 'error': ''}
        
        # === 验证阶段 ===
        validation_errors = []
        
        # 验证所有设备存在且未分配
        for update in device_updates:
            device_id = update['device_id']
            rack_u = update['rack_u']
            
            device = find_instance_by_id(device_id)
            if not device:
                validation_errors.append(f'Device {device_id} not found')
                continue
                
            if device.get('rack_id'):
                validation_errors.append(f'Device {device_id} already assigned to rack {device["rack_id"]}')
                continue
                
            if not isinstance(rack_u, int) or rack_u <= 0:
                validation_errors.append(f'Invalid rack_u {rack_u} for device {device_id}')
                continue
                
            device_type = next((dt for dt in device_types if dt['name'] == device['device_type']), None)
            if device_type:
                device_height = device_type.get('height_u', 1)
                if rack_u + device_height - 1 > rack['height_u']:
                    validation_errors.append(f'Device {device_id} exceeds rack height at U{rack_u}')
                    
        if validation_errors:
            return {'success': False, 'device_count': 0, 'error': '; '.join(validation_errors)}
        
        # === 空间冲突检查 ===
        devices_in_rack = [inst for inst in device_instances if inst.get('rack_id') == rack_id]
        conflict_errors = []
        occupied_positions = set()
        
        # 添加现有设备占用的位置
        for existing_device in devices_in_rack:
            if existing_device.get('rack_u'):
                existing_device_type = next((dt for dt in device_types if dt['name'] == existing_device['device_type']), None)
                existing_height = existing_device_type.get('height_u', 1) if existing_device_type else 1
                existing_start = existing_device['rack_u']
                
                for u in range(existing_start, existing_start + existing_height):
                    occupied_positions.add(u)
        
        # 验证批次内无冲突
        batch_positions = set()
        for update in device_updates:
            device_id = update['device_id']
            rack_u = update['rack_u']
            device = find_instance_by_id(device_id)
            device_type = next((dt for dt in device_types if dt['name'] == device['device_type']), None)
            device_height = device_type.get('height_u', 1) if device_type else 1
            
            device_positions = set(range(rack_u, rack_u + device_height))
            
            if device_positions & occupied_positions:
                conflict_errors.append(f'Device {device_id} (U{rack_u}-{rack_u + device_height - 1}) conflicts with existing devices')
                continue
                
            if device_positions & batch_positions:
                conflict_errors.append(f'Device {device_id} (U{rack_u}-{rack_u + device_height - 1}) conflicts with other devices in this batch')
                continue
                
            batch_positions.update(device_positions)
                    
        if conflict_errors:
            return {'success': False, 'device_count': 0, 'error': '; '.join(conflict_errors)}
        
        # === 执行上架 ===
        original_device_states = {}
        for update in device_updates:
            device = find_instance_by_id(update['device_id'])
            original_device_states[update['device_id']] = {
                'rack_id': device.get('rack_id'),
                'rack_u': device.get('rack_u')
            }
        
        original_rack_devices = list(rack.get('devices', []))
        
        try:
            # 更新所有设备
            updated_device_ids = []
            for update in device_updates:
                device_id = update['device_id']
                rack_u = update['rack_u']
                device = find_instance_by_id(device_id)
                
                device['rack_id'] = rack_id
                device['rack_u'] = rack_u
                updated_device_ids.append(device_id)
            
            # 更新机柜设备列表
            if 'devices' not in rack:
                rack['devices'] = []
            rack['devices'].extend(updated_device_ids)
            
            # 保存数据
            save_data(DEVICE_INSTANCES_FILE, device_instances)
            save_data(RACKS_FILE, racks)
            
            return {'success': True, 'device_count': len(updated_device_ids), 'error': ''}
            
        except Exception as save_error:
            # 回滚操作
            for device_id, original_state in original_device_states.items():
                device = find_instance_by_id(device_id)
                device['rack_id'] = original_state['rack_id']
                device['rack_u'] = original_state['rack_u']
            
            rack['devices'] = original_rack_devices
            
            return {'success': False, 'device_count': 0, 'error': f'Save failed: {str(save_error)}'}
        
    except Exception as e:
        return {'success': False, 'device_count': 0, 'error': str(e)}

# --- Data Consistency Check and Repair API ---
@app.route('/check_data_consistency', methods=['POST'])
def check_data_consistency():
    try:
        load_all_data_before_request()
        
        inconsistencies = []
        repairs_made = []
        
        # Check 1: Devices in rack device lists but with null rack_id
        for rack in racks:
            if 'devices' in rack:
                devices_to_remove = []
                for device_id in rack['devices']:
                    device = find_instance_by_id(device_id)
                    if not device:
                        inconsistencies.append(f"Rack {rack['name']} contains non-existent device {device_id}")
                        devices_to_remove.append(device_id)
                    elif device.get('rack_id') != rack['id']:
                        inconsistencies.append(f"Device {device_id} in rack {rack['name']} but has rack_id: {device.get('rack_id')}")
                        devices_to_remove.append(device_id)
                
                # Auto-repair: Remove inconsistent devices from rack list
                if devices_to_remove:
                    for device_id in devices_to_remove:
                        rack['devices'].remove(device_id)
                    repairs_made.append(f"Removed {len(devices_to_remove)} inconsistent devices from rack {rack['name']}")
        
        # Check 2: Devices with rack_id but not in rack device list
        for device in device_instances:
            if device.get('rack_id'):
                rack = next((r for r in racks if r['id'] == device['rack_id']), None)
                if not rack:
                    inconsistencies.append(f"Device {device['id']} has invalid rack_id: {device['rack_id']}")
                    # Auto-repair: Clear invalid rack_id
                    device['rack_id'] = None
                    device['rack_u'] = None
                    repairs_made.append(f"Cleared invalid rack_id for device {device['id']}")
                elif 'devices' not in rack or device['id'] not in rack['devices']:
                    inconsistencies.append(f"Device {device['id']} has rack_id {device['rack_id']} but not in rack device list")
                    # Auto-repair: Add device to rack list
                    if 'devices' not in rack:
                        rack['devices'] = []
                    rack['devices'].append(device['id'])
                    repairs_made.append(f"Added device {device['id']} to rack {rack['name']} device list")
        
        # Check 3: Duplicate devices in rack lists
        for rack in racks:
            if 'devices' in rack:
                unique_devices = list(set(rack['devices']))
                if len(unique_devices) != len(rack['devices']):
                    duplicates = len(rack['devices']) - len(unique_devices)
                    inconsistencies.append(f"Rack {rack['name']} has {duplicates} duplicate device entries")
                    # Auto-repair: Remove duplicates
                    rack['devices'] = unique_devices
                    repairs_made.append(f"Removed {duplicates} duplicate devices from rack {rack['name']}")
        
        # Check 4: Duplicate device instance_name
        name_to_first_id = {}
        duplicate_ids = []
        for dev in device_instances:
            name = dev.get('instance_name')
            if name in name_to_first_id:
                duplicate_ids.append(dev['id'])
            else:
                name_to_first_id[name] = dev['id']
        if duplicate_ids:
            inconsistencies.append(f'Found {len(duplicate_ids)} duplicate device instances with same instance_name')
            # Auto-repair: delete duplicate device instances and clean up racks
            device_instances[:] = [dev for dev in device_instances if dev['id'] not in duplicate_ids]
            for rack in racks:
                if 'devices' in rack:
                    rack['devices'] = [d for d in rack['devices'] if d not in duplicate_ids]
            repairs_made.append(f'Removed {len(duplicate_ids)} duplicate device instances (duplicates by instance_name)')
        
        # Save repairs if any were made
        if repairs_made:
            save_data(DEVICE_INSTANCES_FILE, device_instances)
            save_data(RACKS_FILE, racks)
        
        return jsonify({
            'message': 'Data consistency check completed',
            'total_racks': len(racks),
            'total_devices': len(device_instances),
            'inconsistencies_found': len(inconsistencies),
            'repairs_made': len(repairs_made),
            'issues_found': len(inconsistencies),  # 别名，保持兼容性
            'auto_repairs': len(repairs_made),     # 别名，保持兼容性
            'details': {
                'inconsistencies': inconsistencies,
                'repairs': repairs_made
            }
        }), 200
        
    except Exception as e:
        print(f'Error in data consistency check: {e}')
        return jsonify({'message': 'Internal server error during consistency check'}), 500

# --- Batch Rack Unmount API ---
@app.route('/batch_rack_unmount', methods=['POST'])
def batch_rack_unmount():
    try:
        load_all_data_before_request()
        data = request.get_json()
        
        if not data:
            return jsonify({'message': 'Invalid JSON data received'}), 400
            
        device_ids = data.get('device_ids', [])  # List of device IDs to unmount
        unmount_type = data.get('unmount_type', 'selected')  # 'selected', 'rack_all'
        rack_id = data.get('rack_id', None)  # For rack_all type
        
        if not device_ids and unmount_type != 'rack_all':
            return jsonify({'message': 'No devices specified for unmounting'}), 400
            
        if unmount_type == 'rack_all' and not rack_id:
            return jsonify({'message': 'rack_id required for rack_all unmount type'}), 400
            
        # Determine devices to unmount
        devices_to_unmount = []
        
        if unmount_type == 'rack_all':
            # Find all devices in the specified rack
            devices_to_unmount = [device for device in device_instances if device.get('rack_id') == rack_id]
        else:
            # Use specified device IDs
            devices_to_unmount = [device for device in device_instances if device['id'] in device_ids]
            
        if not devices_to_unmount:
            return jsonify({'message': 'No devices found to unmount'}), 404
            
        # Validate devices are powered off before unmounting
        validation_errors = []
        for device in devices_to_unmount:
            if not device.get('rack_id'):
                validation_errors.append(f'Device {device.get("instance_name", device["id"])} is already unmounted.')
            if device.get('power_status') == 'on':
                validation_errors.append(f'设备 {device.get("instance_name", device["id"])} 必须先下电才能下架。')
                
        if validation_errors:
            return jsonify({'message': '验证失败，操作已取消。', 'errors': validation_errors}), 400
            
        # Perform batch unmounting
        success_count = 0
        unmounted_devices = []
        racks_to_update = set()
        
        for device in devices_to_unmount:
            device_rack_id = device.get('rack_id')
            if device_rack_id:
                # Remove from rack's device list
                rack = next((r for r in racks if r['id'] == device_rack_id), None)
                if rack and 'devices' in rack and device['id'] in rack['devices']:
                    rack['devices'].remove(device['id'])
                    racks_to_update.add(device_rack_id)
                
                # Clear device's rack information
                device['rack_id'] = None
                device['rack_u'] = None
                
                unmounted_devices.append({
                    'device_id': device['id'],
                    'device_name': device['instance_name'],
                    'previous_rack_id': device_rack_id
                })
                success_count += 1
                
        # Save data atomically
        save_data(DEVICE_INSTANCES_FILE, device_instances)
        save_data(RACKS_FILE, racks)
        
        print(f'[API] batch_unmount: {success_count} devices unmounted. Type: {unmount_type}.')
        return jsonify({
            'message': f'Batch unmount successful',
            'success_count': success_count,
            'unmounted_devices': unmounted_devices,
            'affected_racks': list(racks_to_update)
        }), 200
        
    except Exception as e:
        print(f'Error in batch rack unmount: {e}')
        return jsonify({'message': 'Internal server error during batch rack unmount'}), 500

@app.route('/batch_add_racks', methods=['POST'])
def batch_add_racks():
    """批量添加机柜"""
    global racks  # 在函数开始就声明全局变量
    try:
        load_all_data_before_request()
        data = request.get_json()
        
        # 验证必需字段
        required_fields = ['room_id', 'rack_row', 'start_number', 'end_number']
        for field in required_fields:
            if field not in data:
                return jsonify({'success': False, 'message': f'缺少必需字段: {field}'}), 400
        
        room_id = data['room_id']
        rack_row = data['rack_row'].strip()  # 支持任意自定义排名
        start_number = int(data['start_number'])
        end_number = int(data['end_number'])
        height_u = int(data.get('height_u', 42))  # 默认42U
        department = data.get('department', '').strip()
        
        # 验证输入
        if not rack_row:
            return jsonify({'success': False, 'message': '机柜排名不能为空'}), 400
        
        # 限制排名长度，避免生成过长的机柜名
        if len(rack_row) > 10:
            return jsonify({'success': False, 'message': '机柜排名长度不能超过10个字符'}), 400
        
        if start_number > end_number:
            return jsonify({'success': False, 'message': '开始编号不能大于结束编号'}), 400
        
        if start_number < 1 or end_number > 999:
            return jsonify({'success': False, 'message': '编号范围应在1-999之间'}), 400
        
        if height_u < 1 or height_u > 100:
            return jsonify({'success': False, 'message': '机柜高度应在1-100U之间'}), 400
        
        # 一次批量创建数量限制
        if end_number - start_number + 1 > 100:
            return jsonify({'success': False, 'message': '一次最多只能创建100个机柜'}), 400
        
        # 验证机房是否存在
        if not any(room['id'] == room_id for room in rooms):
            return jsonify({'success': False, 'message': '指定的机房不存在'}), 400
        
        # 获取同一机房内现有机柜名称（只检查同机房重复）
        existing_rack_names_in_room = {rack['name'] for rack in racks if rack['room_id'] == room_id}
        
        # 生成要创建的机柜列表
        racks_to_create = []
        duplicates = []
        
        for number in range(start_number, end_number + 1):
            # 生成机柜名称：支持任意排名，如A01, B02, North10, Zone-A15 等
            rack_name = f"{rack_row}{number:02d}"
            
            # 只检查同机房内是否重名
            if rack_name in existing_rack_names_in_room:
                duplicates.append(rack_name)
                continue
            
            rack_id = f"rack_{int(time.time() * 1000)}_{number}"
            new_rack = {
                'id': rack_id,
                'name': rack_name,
                'room_id': room_id,
                'height_u': height_u,
                'department': department,
                'devices': []
            }
            racks_to_create.append(new_rack)
            existing_rack_names_in_room.add(rack_name)  # 避免在同一批次中重复
            time.sleep(0.001)  # 确保ID唯一性
        
        if not racks_to_create and duplicates:
            return jsonify({
                'success': False, 
                'message': f'所有机柜名称在该机房中都已存在: {", ".join(duplicates)}'
            }), 400
        
        # 添加新机柜到全局数据
        racks.extend(racks_to_create)
        save_data(RACKS_FILE, racks)
        
        # 准备响应信息
        created_count = len(racks_to_create)
        room_name = next((room['name'] for room in rooms if room['id'] == room_id), '未知机房')
        
        response_data = {
            'success': True,
            'message': f'批量添加机柜成功',
            'created_count': created_count,
            'created_racks': [rack['name'] for rack in racks_to_create],
            'room_name': room_name
        }
        
        if duplicates:
            response_data['duplicates'] = duplicates
            response_data['message'] += f'，成功创建 {created_count} 个机柜'
            if len(duplicates) > 0:
                response_data['message'] += f'，跳过 {len(duplicates)} 个在该机房中重名的机柜'
        
        print(f"[API] batch_add_racks: In room '{room_name}', created {created_count} racks ('{rack_row}' from {start_number} to {end_number}). Skipped {len(duplicates)} duplicates.")
        return jsonify(response_data)
        
    except ValueError as e:
        print(f"批量添加机柜时数据格式错误: {e}")
        return jsonify({'success': False, 'message': f'数据格式错误: {str(e)}'}), 400
    except Exception as e:
        print(f"[API ERROR] batch_add_racks: {e}")
        return jsonify({'success': False, 'message': f'批量添加机柜失败: {str(e)}'}), 500

@app.route('/update_rack/<rack_id>', methods=['PUT'])
def update_rack(rack_id):
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
                return jsonify({'message': '机柜高度必须是1-100之间的正整数'}), 400
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

        # Find and update the rack
        rack_found = False
        for rack in racks:
            if rack['id'] == rack_id:
                # Check if the new name conflicts with other racks in the same room
                if name != rack['name']:  # Only check if name is being changed
                    for other_rack in racks:
                        if other_rack['id'] != rack_id and other_rack['room_id'] == room_id and other_rack['name'] == name:
                            return jsonify({'message': f'机房中已存在名为 {name} 的机柜'}), 400

                rack['name'] = name
                rack['room_id'] = room_id
                rack['height_u'] = height_u
                rack['department'] = department
                rack_found = True
                break

        if not rack_found:
            return jsonify({'message': '机柜不存在'}), 404

        # Save changes to file
        save_data(RACKS_FILE, racks)
        return jsonify({'message': '机柜更新成功'})

    except Exception as e:
        print(f"Error updating rack: {str(e)}")
        return jsonify({'message': '更新机柜失败'}), 500



@app.route('/batch_delete_racks', methods=['POST'])
def batch_delete_racks():
    try:
        data = request.get_json()
        rack_ids_to_delete = data.get('rack_ids', [])

        if not isinstance(rack_ids_to_delete, list) or not rack_ids_to_delete:
            return jsonify({'message': '请提供一个包含机柜ID的列表。'}), 400

        global racks, device_instances
        
        original_rack_count = len(racks)
        original_device_count = len(device_instances)

        # Get a set for faster lookups
        rack_ids_set = set(rack_ids_to_delete)

        # Filter out devices that are in the racks to be deleted
        device_instances = [device for device in device_instances if device.get('rack_id') not in rack_ids_set]
        
        # Filter out the racks to be deleted
        racks = [rack for rack in racks if rack.get('id') not in rack_ids_set]

        deleted_rack_count = original_rack_count - len(racks)
        deleted_device_count = original_device_count - len(device_instances)
        
        if deleted_rack_count == 0:
             return jsonify({'message': '没有找到任何要删除的机柜。可能ID不正确。'}), 404

        # Save changes to files
        save_data(RACKS_FILE, racks)
        save_data(DEVICE_INSTANCES_FILE, device_instances)

        message = f'成功删除 {deleted_rack_count} 个机柜。'
        if deleted_device_count > 0:
            message += f' 同时删除了关联的 {deleted_device_count} 个设备实例。'

        return jsonify({'message': message, 'success': True})

    except Exception as e:
        print(f"Error during batch rack deletion: {str(e)}")
        return jsonify({'message': '批量删除机柜时发生内部错误。'}), 500

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
            return jsonify({'message': '设备实例名称不能为空。'}), 400

        # Check for duplicate names
        for instance in device_instances:
            if instance['id'] != instance_id and instance['instance_name'] == new_name:
                return jsonify({'message': f'已存在名为 "{new_name}" 的设备实例。'}), 400

        # Find and update the instance
        instance_found = False
        for instance in device_instances:
            if instance['id'] == instance_id:
                instance['instance_name'] = new_name
                instance_found = True
                break

        if not instance_found:
            return jsonify({'message': '设备实例未找到。'}), 404

        save_data(DEVICE_INSTANCES_FILE, device_instances)
        return jsonify({'message': '设备实例更新成功。'}), 200

    except Exception as e:
        print(f"Error updating device instance: {str(e)}")
        return jsonify({'message': '更新设备实例时发生内部错误。'}), 500

@app.route('/devices/batch_power', methods=['POST'])
def batch_power_devices():
    """Batch power on or off devices"""
    try:
        data = request.get_json()
        device_ids = data.get('device_ids', [])
        status = data.get('status') # 'on' or 'off'

        if not device_ids or status not in ['on', 'off']:
            return jsonify({'message': '缺少 device_ids 或无效的状态。'}), 400

        success_count = 0
        error_details = []
        
        for device_id in device_ids:
            instance = find_instance_by_id(device_id)
            if not instance:
                error_details.append(f"设备 {device_id} 未找到。")
                continue

            # Powering on requires the device to be mounted
            if status == 'on' and not instance.get('rack_id'):
                error_details.append(f"设备 {instance.get('instance_name')} 未上架，无法上电。")
                continue
            
            instance['power_status'] = status
            success_count += 1
        
        save_data(DEVICE_INSTANCES_FILE, device_instances)

        message = f"成功 {status_map.get(status, '')} {success_count} 个设备。"
        if error_details:
             message += f" {len(error_details)} 个设备操作失败。"

        return jsonify({
            'message': message,
            'success_count': success_count,
            'errors': error_details
        }), 200 if not error_details else 206

    except Exception as e:
        print(f"Error in batch power operation: {str(e)}")
        return jsonify({'message': '批量电源操作时发生内部错误。'}), 500

status_map = {
    'on': '开启',
    'off': '关闭'
}

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
                error_details.append(f'设备 "{instance.get("instance_name")}" 必须先下架才能删除。')
        
        if error_details:
            return jsonify({'message': '验证失败，部分设备仍处于上架状态。', 'errors': error_details}), 400

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

        return jsonify({'message': f'成功删除 {deleted_count} 个设备。'}), 200

    except Exception as e:
        print(f"Error during batch device deletion: {e}")
        return jsonify({'message': '批量删除期间发生内部错误。'}), 500

if __name__ == '__main__':
    app.run(host='172.31.60.204', port=58000, debug=True) 