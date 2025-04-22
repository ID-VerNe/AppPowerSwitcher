import logging
import configparser
import os
import sys

# 获取当前模块的 logger
logger = logging.getLogger(__name__)

# --- 常量定义 ---
# 配置文件的目录和名称
CONFIG_DIR_NAME = "config"
CONFIG_FILE_NAME = "app_config.ini"

# 获取项目根目录
# 此文件位于 [project_root]/src/infrastructure/configuration/
# 需要向上三级目录到达 project_root/
# dirname(__file__)          = src/infrastructure/configuration/
# dirname(dirname(__file__))   = src/infrastructure/
# dirname(dirname(dirname(__file__))) = src/
# dirname(dirname(dirname(dirname(__file__)))) = project_root/
PROJECT_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

# 配置文件的完整路径
CONFIG_FILE_PATH = os.path.join(PROJECT_ROOT, CONFIG_DIR_NAME, CONFIG_FILE_NAME)

# 配置文件 section 名称
SECTION_GENERAL = "General"
SECTION_PROCESS_POWER_MAP = "ProcessPowerMap"

# General section keys
KEY_DEFAULT_POWER_PLAN = "default_power_plan"
KEY_LOG_LEVEL = "log_level"

# Example common GUIDs (for default config file creation)
# Note: These are common but may vary slightly; user should verify with 'powercfg /list'
GUID_BALANCED = "381b4222-f694-41f0-9685-ff5bb260df2e"
GUID_HIGH_PERFORMANCE = "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
GUID_POWER_SAVER = "a1841308-3541-4fab-bc81-f71556f20b4a"

# --- ConfigManager Class ---

class ConfigManager:
    """
    管理应用程序的配置，包括从文件加载和保存，以及提供配置数据的访问接口。
    使用 .ini 文件格式。电源计划预期为 GUID。
    """
    def __init__(self, config_file_path=CONFIG_FILE_PATH):
        """
        初始化 ConfigManager.

        Args:
            config_file_path: 配置文件的完整路径。默认为项目根目录下的 config/app_config.ini。
        """
        self._config_file_path = config_file_path
        self._config_parser = configparser.ConfigParser()
        # Internal storage for parsed config data, mainly the process-power map
        # Expect process name (lowercase) -> power plan GUID string
        self._app_power_map = {}
        # Expect default power plan GUID string
        self._default_power_plan = None
        # Expect log level string, e.g., "INFO", "DEBUG"
        self._log_level = None

        logger.debug(f"ConfigManager initialized with config file path: {self._config_file_path}")

    def load_config(self):
        """
        从配置文件加载应用程序配置。
        电源计划预期为 GUID。
        """
        logger.info(f"Attempting to load configuration from: {self._config_file_path}")
        if not os.path.exists(self._config_file_path):
            logger.warning(f"Configuration file not found at {self._config_file_path}. Creating with default structure.")
            self._create_default_config_file()
            # After creating, attempt to load again. If creation failed, this load will fail too.
            if not os.path.exists(self._config_file_path):
                 logger.error(f"Failed to create default config file at {self._config_file_path}. Cannot load config.")
                 return False

        try:
            # Read the configuration file
            # configparser.read() returns a list of filenames successfully read.
            read_files = self._config_parser.read(self._config_file_path, encoding='utf-8')
            if not read_files:
                # This might happen if the file exists but is empty or read fails silently
                logger.error(f"ConfigParser failed to read file: {self._config_file_path}. Read files: {read_files}")
                # Attempt to reinitialize internal state with defaults or empty state on read failure
                self._default_power_plan = GUID_BALANCED
                self._log_level = "INFO"
                self._app_power_map = {}
                return False

            logger.info("Configuration file read successfully.")

            # --- Parse General Section ---
            if self._config_parser.has_section(SECTION_GENERAL):
                try:
                    # Get default power plan GUID (strip whitespace)
                    # Use a common default GUID as fallback if not found in config or section missing/empty
                    self._default_power_plan = self._config_parser.get(SECTION_GENERAL, KEY_DEFAULT_POWER_PLAN, fallback=GUID_BALANCED).strip()
                    logger.info(f"Loaded default power plan GUID: {self._default_power_plan}")

                    # Get log level string (strip whitespace, convert to uppercase)
                    self._log_level = self._config_parser.get(SECTION_GENERAL, KEY_LOG_LEVEL, fallback="INFO").strip().upper()
                    valid_levels = ['CRITICAL', 'ERROR', 'WARNING', 'INFO', 'DEBUG', 'NOTSET']
                    if self._log_level not in valid_levels:
                        logger.warning(f"Invalid log level '{self._log_level}' in config '{KEY_LOG_LEVEL}'. Falling back to 'INFO'. Valid levels are: {', '.join(valid_levels)}")
                        self._log_level = "INFO"
                    logger.info(f"Loaded log level setting: {self._log_level}")

                except Exception as e:
                    logger.error(f"Error parsing '{SECTION_GENERAL}' section from '{self._config_file_path}': {e}", exc_info=True)
                    # Ensure internal state is set to defaults on error in this section
                    self._default_power_plan = GUID_BALANCED
                    self._log_level = "INFO"

            else:
                logger.warning(f"'{SECTION_GENERAL}' section not found in config file. Using default General settings.")
                # Ensure internal state is set to defaults
                self._default_power_plan = GUID_BALANCED
                self._log_level = "INFO"

            # --- Parse ProcessPowerMap Section ---
            self._app_power_map = {} # Clear previous map before parsing
            if self._config_parser.has_section(SECTION_PROCESS_POWER_MAP):
                try:
                    # Parse all items in the map section. Expect process name and power plan GUID.
                    for process_name, power_plan_guid_str in self._config_parser.items(SECTION_PROCESS_POWER_MAP):
                        clean_process_name = process_name.strip()
                        clean_power_plan_guid = power_plan_guid_str.strip()
                        # Basic validation: check if it's a process name and something that looks like a GUID
                        # A more robust check could use a regex for GUID format.
                        # For now, just check if both key and value are non-empty after stripping.
                        if clean_process_name and clean_power_plan_guid:
                            # Store process names lowercase for case-insensitive matching later
                            # Store GUID as cleaned string. Validation of GUID format happens when powercfg is called.
                            self._app_power_map[clean_process_name.lower()] = clean_power_plan_guid
                            # logger.debug(f"Loaded map entry: '{clean_process_name}' -> '{clean_power_plan_guid}'") # Too verbose potentially
                        # else: commented entries or entries with empty key/value will be skipped
                        # logger.debug(f"Skipping potential invalid map entry (empty key/value or comment): '{process_name}' -> '{power_plan_guid_str}'")

                    logger.info(f"Loaded {len(self._app_power_map)} process-to-power plan map entries (expecting GUIDs).")

                except Exception as e:
                    logger.error(f"Error parsing '{SECTION_PROCESS_POWER_MAP}' section from '{self._config_file_path}': {e}", exc_info=True)
                    self._app_power_map = {} # Clear map on failure

            else:
                logger.warning(f"'{SECTION_PROCESS_POWER_MAP}' section not found in config file. Process-Power map is empty.")
                self._app_power_map = {}

            logger.info("Configuration loading finished.")
            return True

        except configparser.Error as e:
            logger.error(f"ConfigParser error while loading file '{self._config_file_path}': {e}", exc_info=True)
            # Ensure internal state defaults on major configparser error
            self._default_power_plan = GUID_BALANCED
            self._log_level = "INFO"
            self._app_power_map = {}
            return False
        except Exception as e:
            logger.error(f"An unexpected error occurred during config loading from '{self._config_file_path}': {e}", exc_info=True)
            # Ensure internal state defaults on any exception
            self._default_power_plan = GUID_BALANCED
            self._log_level = "INFO"
            self._app_power_map = {}
            return False

    def save_config(self, config_data: dict):
        """
        将配置数据保存到配置文件。
        期望 config_data 中的电源计划值为 GUID 字符串。

        Args:
            config_data: 包含要保存的配置数据的字典。
                         格式示例 (键和值应为字符串):
                         {
                             "General": {"default_power_plan": "guid", "log_level": "..."},
                             "ProcessPowerMap": {"process1.exe": "guid1", "process2.exe": "guid2"}
                         }
        Returns:
            True if saved successfully, False otherwise.
        """
        logger.info(f"Attempting to save configuration to: {self._config_file_path}")

        # Create a new ConfigParser instance and rebuild from the data provided
        # This ensures we don't save back comments or format from the old file unless we explicitly handle them.
        self._config_parser = configparser.ConfigParser()

        try:
            # Populate parser from the input dictionary
            for section_name, section_data in config_data.items():
                # Ensure section data is a dictionary
                if not isinstance(section_data, dict):
                     logger.warning(f"Skipping save for section '{section_name}', expected dictionary but got {type(section_data).__name__}")
                     continue

                # Add section if it doesn't exist
                if section_name not in self._config_parser:
                     self._config_parser.add_section(section_name)

                # Add key-value pairs to the section
                for key, value in section_data.items():
                    # Ensure keys/values are strings and strip whitespace from values
                    cleaned_key = str(key).strip()
                    cleaned_value = str(value).strip()
                    if cleaned_key: # Only save non-empty keys
                        self._config_parser.set(section_name, cleaned_key, cleaned_value)
                    else:
                         logger.warning(f"Skipping save entry with empty key in section '{section_name}': '{key}' -> '{value}'")

            # Ensure config directory exists
            config_dir = os.path.dirname(self._config_file_path)
            # Check if config_dir is an empty string (can happen if path is just a filename)
            if config_dir and not os.path.exists(config_dir):
                 logger.debug(f"Config directory '{config_dir}' not found, creating it.")
                 try:
                     os.makedirs(config_dir, exist_ok=True)
                 except OSError as e:
                      logger.error(f"Failed to create config directory '{config_dir}': {e}", exc_info=True)
                      return False # Cannot save if directory creation fails

            # Write to the file
            try:
                with open(self._config_file_path, 'w', encoding='utf-8') as configfile:
                    self._config_parser.write(configfile)
                logger.info("Configuration saved successfully.")
            except IOError as e:
                 logger.error(f"Failed to write config file '{self._config_file_path}': {e}", exc_info=True)
                 return False

            # After saving, update internal state by loading back from file.
            # This ensures the internal state reflects exactly what's on disk
            # and handles any default value fallbacks or format changes configparser might introduce on read.
            # Return the success status of the reload.
            reload_success = self.load_config()
            if not reload_success:
                 logger.error("Failed to reload config after saving. Internal state might not match file.")
            return reload_success # Return status based on whether saving AND subsequent loading were successful

        except Exception as e:
            logger.error(f"An unexpected error occurred during config saving to '{self._config_file_path}': {e}", exc_info=True)
            return False

    def _create_default_config_file(self):
        """
        创建默认配置文件的内容和结构。
        使用默认 GUID 作为示例。
        """
        logger.info(f"Creating default config file at {self._config_file_path}")
        config = configparser.ConfigParser()

        # Add sections and default values
        config[SECTION_GENERAL] = {
            # Use default GUID for Balanced plan (common, but user should verify with 'powercfg /list')
            KEY_DEFAULT_POWER_PLAN: GUID_BALANCED,
            KEY_LOG_LEVEL: 'INFO',             # Default logging level
        }

        config[SECTION_PROCESS_POWER_MAP] = {
            # Example entries (commented out) using example GUIDs
            # Get your actual process names and power plan GUIDs using 'powercfg /list' in Command Prompt or PowerShell
            '# Example: chrome.exe': GUID_HIGH_PERFORMANCE, # Example: Chrome to High performance
            '# Example: steam.exe': GUID_HIGH_PERFORMANCE,  # Example: Steam to High performance
            '# Example: vlc.exe': GUID_BALANCED,            # Example: VLC to Balanced
            '# Example: notepad++.exe': GUID_POWER_SAVER,   # Example: Notepad++ to Power saver

            # Add your mappings below (uncomment and replace with your process names and GUIDs):
            # mygame.exe = 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c # Example: Your Game to High performance
        }

        # Ensure config directory exists
        config_dir = os.path.dirname(self._config_file_path)
         # Check if config_dir is not empty before attempting creation
        if config_dir and not os.path.exists(config_dir):
             logger.debug(f"Config directory '{config_dir}' not found during default creation, creating it.")
             try:
                 os.makedirs(config_dir, exist_ok=True)
             except OSError as e:
                  logger.error(f"Failed to create config directory '{config_dir}' during default file creation: {e}", exc_info=True)
                  return # Abort default file creation if directory creation fails

        try:
            with open(self._config_file_path, 'w', encoding='utf-8') as configfile:
                config.write(configfile)
            logger.info("Default config file created successfully.")
        except IOError as e:
            logger.error(f"Error writing default config file at '{self._config_file_path}': {e}", exc_info=True)
        except Exception as e:
             logger.error(f"An unexpected error occurred during default config file creation at '{self._config_file_path}': {e}", exc_info=True)

    def get_power_plan_for_process(self, process_name: str):
        """
        根据进程名称查询对应的电源计划 GUID。
        期望配置文件中的值为 GUID。

        Args:
            process_name: 前台应用的进程名称 (例如 "chrome.exe")。

        Returns:
            查找到的电源计划 GUID 字符串。
            如果进程未在映射中，返回 None。
            返回的 GUID 未经校验，仅是配置文件中的字符串。
        """
        if not process_name:
            return None
        # Use case-insensitive lookup because process names from OS might vary in casing
        lookup_name = process_name.lower().strip()
        power_plan_guid = self._app_power_map.get(lookup_name)
        # logger.debug(f"Looking up power plan for process '{process_name}' ({lookup_name}) -> '{power_plan_guid}'") # Too verbose
        return power_plan_guid

    def get_default_power_plan(self):
        """
        获取配置文件中定义的默认电源计划 GUID。
        """
        return self._default_power_plan

    def get_log_level(self):
        """
        获取配置文件中定义的日志级别字符串 (例如 "INFO").
        """
        return self._log_level

    def get_app_power_map(self):
        """
        获取当前加载的应用进程到电源计划 GUID 的映射字典。
        返回的数据是一个副本，防止外部直接修改内部状态。
        """
        return self._app_power_map.copy()

    # Methods for updating internal state from external data (e.g., from GUI)
    # These methods are intended for GUI interaction and update internal state, NOT the config file.
    # Call save_config AFTER using these if you want changes to persist.

    def update_app_power_map(self, new_map_data: dict):
        """
        更新内部的应用进程到电源计划 GUID 的映射数据。
        通常在从 GUI 获取用户修改的配置后调用。
        请注意，调用此方法不会自动保存到文件，需要显式调用 save_config。
        期望输入字典的键为进程名字符串，值为 GUID 字符串。

        Args:
            new_map_data: 新的映射字典，例如 {"process1.exe": "guid1", "process2.exe": "guid2"}.
        """
        if not isinstance(new_map_data, dict):
             logger.error("update_app_power_map expects a dictionary.")
             return

        # Build the new map from the input data, ensuring keys are lowercase and values are stripped strings
        self._app_power_map = {
            str(k).strip().lower(): str(v).strip()
            for k, v in new_map_data.items()
            # Only include entries where both key (process name) and value (GUID) are non-empty after stripping
            if str(k).strip() and str(v).strip()
        }
        logger.info(f"Internal app power map updated. Contains {len(self._app_power_map)} entries (expecting GUIDs).")

    def update_general_settings(self, default_power_plan_guid=None, log_level=None):
        """
        更新内部的通用设置数据。
        默认电源计划预期为 GUID 字符串。
        通常在从 GUI 获取用户修改的配置后调用。
        请注意，调用此方法不会自动保存到文件，需要显式调用 save_config。

        Args:
            default_power_plan_guid: 新的默认电源计划 GUID 字符串。
            log_level: 新的日志级别字符串 (例如 "INFO").
        """
        if default_power_plan_guid is not None:
             self._default_power_plan = str(default_power_plan_guid).strip()
             logger.info(f"Internal default power plan GUID updated to: {self._default_power_plan}")

        if log_level is not None:
             log_level_str = str(log_level).strip().upper()
             valid_levels = ['CRITICAL', 'ERROR', 'WARNING', 'INFO', 'DEBUG', 'NOTSET']
             if log_level_str not in valid_levels:
                  logger.warning(f"Invalid log level '{log_level_str}' provided for update. Not updating log level.")
             else:
                self._log_level = log_level_str
                logger.info(f"Internal log level updated to: {self._log_level}")
