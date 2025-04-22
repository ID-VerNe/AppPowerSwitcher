import logging
import os
from datetime import datetime

# 项目根目录，便于日志文件存放
BASE_DIR = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

# 日志文件存放目录
LOG_DIR = os.path.join(BASE_DIR, "logs")
# 确保日志目录存在
if not os.path.exists(LOG_DIR):
    os.makedirs(LOG_DIR)

# 日志文件名称 (可以考虑使用日期或时间命名，这里简化处理)
LOG_FILE_NAME = "app_power_switcher.log"
LOG_FILE_PATH = os.path.join(LOG_DIR, LOG_FILE_NAME)

# 日志格式 (可以根据需要配置更详细的信息)
# 例如: '[%(asctime)s] [%(levelname)s] [%(name)s] [%(funcName)s:%(lineno)d] - %(message)s'
LOG_FORMAT = '[%(asctime)s] p%(process)d t%(thread)d [%(levelname)s] [%(name)s.%(funcName)s] - %(message)s'
DATE_FORMAT = '%Y-%m-%d %H:%M:%S'

# 默认日志级别 (可以配置)
# logging.DEBUG, logging.INFO, logging.WARNING, logging.ERROR, logging.CRITICAL
DEFAULT_LOG_LEVEL = logging.INFO

def configure_logging(log_level=None):
    """
    配置应用程序的日志系统.

    Args:
        log_level: 可选参数，用于覆盖默认的日志级别。
                   应为 logging 模块定义的级别常量 (如 logging.INFO)。
    """
    # 获取根 logger
    root_logger = logging.getLogger()

    # 避免重复配置 logger
    if root_logger.handlers:
        # 如果已经有 Handlers，说明已经配置过了，直接返回
        # Logging has already been configured, skip
        return

    # 设置日志级别
    level = log_level if log_level is not None else DEFAULT_LOG_LEVEL
    root_logger.setLevel(level)

    # 创建格式器
    formatter = logging.Formatter(LOG_FORMAT, datefmt=DATE_FORMAT)

    # 创建控制台处理器 (Console Handler)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(level) # 控制台输出级别
    console_handler.setFormatter(formatter)

    # 创建文件处理器 (File Handler)
    # 使用 RotatingFileHandler 可以限制日志文件大小和数量
    try:
        from logging.handlers import RotatingFileHandler
        file_handler = RotatingFileHandler(
            LOG_FILE_PATH,
            maxBytes=1024 * 1024 * 5, # 5 MB
            backupCount=5,           # 最多保留 5 个备份文件
            encoding='utf-8'         # 使用 UTF-8 编码
        )
        file_handler.setLevel(level) # 文件输出级别
        file_handler.setFormatter(formatter)
        root_logger.addHandler(file_handler)
        print(f"Logging configured. Output to console and file: {LOG_FILE_PATH}") # 初始提示信息
    except Exception as e:
        # 如果文件处理器创建失败 (例如权限问题), 只配置控制台输出
        print(f"Warning: Could not create file handler at {LOG_FILE_PATH}. Log will only be output to console.")
        print(f"Error details: {e}")

    # 将处理器添加到根 logger
    root_logger.addHandler(console_handler)

    # 获取当前 logger，用于在此模块中记录配置状态
    logger = logging.getLogger(__name__)
    logger.info("Logging system initialized.")
    logger.info(f"Log Level set to: {logging.getLevelName(root_logger.level)}")
    for handler in root_logger.handlers:
         logger.info(f"Handler added: {type(handler).__name__} with level {logging.getLevelName(handler.level)}")

# 可以在程序的入口 (main.py) 调用 configure_logging() 函数来初始化日志系统。
# 例如:
# from src.utils.logging_config import configure_logging
# configure_logging()

# 或者在 main.py 中根据配置加载后设置日志级别:
# from src.utils.logging_config import configure_logging
# from src.infrastructure.configuration.config_manager import ConfigManager
# # ... (加载配置)
# config = ConfigManager()
# config.load_config() # 假设配置中包含了日志级别
# log_level_str = config.get_log_level()
# log_level = getattr(logging, log_level_str.upper(), logging.INFO) # 转换为 logging 级别对象
# configure_logging(log_level)