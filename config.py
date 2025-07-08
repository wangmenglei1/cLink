import os
from dotenv import load_dotenv

# 加载.env文件中的环境变量
load_dotenv()

# 服务器配置
SERVER_HOST = os.getenv('SERVER_HOST', '0.0.0.0')  # 如果环境变量不存在，使用默认值
SERVER_PORT = int(os.getenv('SERVER_PORT', '5000'))  # 转换为整数
DEBUG_MODE = os.getenv('DEBUG_MODE', 'False').lower() == 'true'  # 转换为布尔值

# 数据文件路径
DATA_DIR = os.getenv('DATA_DIR', 'data')
DEVICE_TYPES_FILE = os.path.join(DATA_DIR, 'device_types.json')
DEVICE_INSTANCES_FILE = os.path.join(DATA_DIR, 'device_instances.json')
DEVICE_GROUPS_FILE = os.path.join(DATA_DIR, 'device_groups.json')
CONNECTIONS_FILE = os.path.join(DATA_DIR, 'connections.json')
ROOMS_FILE = os.path.join(DATA_DIR, 'rooms.json')
RACKS_FILE = os.path.join(DATA_DIR, 'racks.json')
GROUP_TEMPLATES_FILE = os.path.join(DATA_DIR, 'group_templates.json')
DEVICE_PORT_TEMPLATES_FILE = os.path.join(DATA_DIR, 'device_port_templates.json')

# 其他配置
MAX_DEVICES_PER_GROUP = int(os.getenv('MAX_DEVICES_PER_GROUP', '2'))
DEFAULT_ENCODING = os.getenv('DEFAULT_ENCODING', 'utf-8') 
