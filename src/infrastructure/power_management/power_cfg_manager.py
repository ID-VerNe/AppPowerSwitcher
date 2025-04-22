import logging
import subprocess
import sys
import shlex # Still useful for potential edge cases, but less critical now
import re

import win32con

# threading is not needed here
# import threading # Ensure threading is not imported

# Get logger for this module
logger = logging.getLogger(__name__)

# --- Constants ---
POWERCARD_COMMAND = "powercfg"
SET_ACTIVE_ARG = "/setactive"
GET_ACTIVE_SCHEME_ARG = "/getactivescheme"
LIST_ARG = "/list" # Used to get all plans

# Regex to parse power scheme information lines from 'powercfg /list' output.
# It looks for:
# - 'GUID:' followed by optional spaces
# - The GUID pattern (captured in group 1)
# - Optional spaces followed by '('
# - The Power Plan Name inside parentheses (captured in group 2)
# - Optional spaces followed by optional '*' (for the active scheme) and more optional spaces
# This regex is designed to be robust against preceding text and language variations, focusing on the GUID and the content in parentheses.
POWER_PLAN_LIST_REGEX = re.compile(r"GUID:\s*([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})\s+\((.*?)\)\s*\*?\s*", re.IGNORECASE)

# Regex to parse the GUID from 'powercfg /getactivescheme' output.
# It looks for 'GUID:' followed by optional spaces and then the GUID pattern (captured in group 1).
# This regex is also robust against preceding text and language variations.
GUID_PARSE_REGEX = re.compile(r"GUID:\s*([0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12})", re.IGNORECASE)

# --- PowerCfgManager Class ---

class PowerCfgManager:
    """
    负责通过 Windows 内置的 powercfg 命令行工具管理电源计划。
    能够加载系统电源方案列表 (GUID 和名称) 并根据 GUID 切换。
    """
    def __init__(self):
        """
        初始化 PowerCfgManager. 加载系统可用电源方案列表。
        """
        logger.debug("Initializing PowerCfgManager instance.")
        # Store mapping from plan name (lowercase) to GUID
        self._name_to_guid_map = {}
        # Store GUID (lowercase) to name map
        self._guid_to_name_map = {}
        # Load schemes during initialization
        self._load_available_schemes()

    def _load_available_schemes(self):
        """
        通过执行 'powercfg /list' 命令加载系统所有可用电源方案的名称和 GUID。
        构建名称到 GUID 和 GUID 到名称 的内部映射。
        这个方法在 __init__ 中调用。
        """
        command = [POWERCARD_COMMAND, LIST_ARG]
        logger.info(f"Loading available power schemes using command: {' '.join(command)}")

        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True, # Decode as text using default system encoding
                check=False, # Do not raise exception for non-zero exit codes
                shell=False,# Do not use shell
                creationflags=win32con.CREATE_NO_WINDOW
            )

            logger.debug(f"PowerCfg list command executed. Return code: {result.returncode}")
            if result.stdout:
                logger.debug(f"PowerCfg list command stdout:\n{result.stdout.strip()}")
            if result.stderr:
                # powercfg /list typically doesn't output to stderr unless there's a major issue
                logger.warning(f"PowerCfg list command stderr:\n{result.stderr.strip()}")

            # On success, parse the output
            if result.returncode == 0:
                lines = result.stdout.splitlines()
                parsed_count = 0
                # Clear existing maps before populating
                self._name_to_guid_map = {}
                self._guid_to_name_map = {}

                for line in lines:
                    match = POWER_PLAN_LIST_REGEX.search(line)
                    if match:
                        # Extract GUID (group 1) and Name (group 2) from regex groups
                        guid = match.group(1).strip()
                        name = match.group(2).strip()
                        if guid and name: # Ensure both GUID and Name were captured
                            # Store mapping using lowercase name for case-insensitive lookup
                            self._name_to_guid_map[name.lower()] = guid
                            # Store mapping using lowercase GUID for case-insensitive lookup
                            self._guid_to_name_map[guid.lower()] = name
                            parsed_count += 1
                            # logger.debug(f"Parsed scheme: '{name}' -> '{guid}'") # Too verbose

                logger.info(f"Successfully loaded {parsed_count} available power schemes.")
                if parsed_count == 0:
                     logger.warning(f"Found 0 power schemes when parsing powercfg '{LIST_ARG}' output. Please check command output format or if any schemes exist.")

            else:
                logger.error(f"PowerCfg list command failed with return code {result.returncode} when loading schemes.", exc_info=True)
                # Clear maps on failure to indicate no schemes were loaded
                self._name_to_guid_map = {}
                self._guid_to_name_map = {}

        except FileNotFoundError:
            logger.error(f"PowerCfg command '{POWERCARD_COMMAND}' not found. Ensure powercfg is in your system's PATH.", exc_info=True)
            # Clear maps on failure
            self._name_to_guid_map = {}
            self._guid_to_name_map = {}
        except Exception as e:
            logger.error(f"An unexpected error occurred while loading available power schemes from powercfg output: {e}", exc_info=True)
            # Clear maps on failure
            self._name_to_guid_map = {}
            self._guid_to_name_map = {}

    def switch_power_plan(self, power_plan_guid: str):
        """
        切换当前的活动电源计划到指定的 GUID.
        此方法直接使用传入的 GUID 调用 powercfg /setactive.

        Args:
            power_plan_guid: 电源计划的 GUID 字符串。

        Returns:
            True if the command executed successfully (return code 0), False otherwise.
            Note: Return code 0 means command syntax was likely correct and executed,
            but doesn't guarantee the OS *actually* switched if the GUID was invalid
            or permission denied in a way powercfg doesn't flag with non-zero.
            Checking get_active_scheme_guid afterwards is more robust.
        """
        if not power_plan_guid:
            logger.warning("Cannot switch power plan - no GUID provided.")
            return False

        # Powercfg /setactive works reliably with GUIDs. Quoting is generally not needed for valid GUIDs.
        # We expect a GUID string here based on the application logic and configuration.
        command_identifier = power_plan_guid.strip() # Use the provided GUID directly after cleaning whitespace

        # Basic check if it looks like a GUID (optional, powercfg will validate anyway)
        # if not re.match(r"^[0-9a-fA-F]{8}(-[0-9a-fA-F]{4}){3}-[0-9a-fA-F]{12}$", command_identifier):
        #      logger.warning(f"Provided power_plan_guid '{power_plan_guid}' does not look like a standard GUID format.")
        #      # Still proceed, powercfg will be the final validator.

        # Build the command. Pass the GUID directly.
        # Use a list of strings for subprocess.run command argument (shell=False is safer)
        command = [POWERCARD_COMMAND, SET_ACTIVE_ARG, command_identifier]
        logger.info(f"Attempting to switch power plan using command: {' '.join(command)}")

        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True, # Decode output as text
                check=False, # Do not raise CalledProcessError on non-zero exit code
                shell=False, # Do not use shell
                creationflags=win32con.CREATE_NO_WINDOW
            )

            logger.debug(f"PowerCfg switch command executed. Return code: {result.returncode}")
            if result.stdout:
                # powercfg /setactive on success usually has no stdout or minimal output if already active
                logger.debug(f"PowerCfg switch command stdout: {result.stdout.strip()}")
            if result.stderr:
                # powercfg /setactive on failure writes error message to stderr
                logger.error(f"PowerCfg switch command stderr: {result.stderr.strip()}")

            if result.returncode == 0:
                # Successful command even if no change needed (powercfg returns 0 if already active or if valid GUID syntax).
                logger.info(f"Successfully sent command to switch power plan to GUID '{power_plan_guid}'.")
                return True
            else:
                # Non-zero return code indicates an issue (invalid GUID, syntax error, permission denied, etc.)
                logger.error(f"PowerCfg command failed with return code {result.returncode} for GUID '{power_plan_guid}'. Check stderr for details.")
                logger.error("Hint: Power plan switching typically requires Administrator privileges or the GUID was invalid.")
                return False

        except FileNotFoundError:
            logger.error(f"PowerCfg command '{POWERCARD_COMMAND}' not found. Ensure powercfg is in your system's PATH.")
            return False
        except Exception as e:
            logger.error(f"An unexpected error occurred while calling powercfg command: {e}", exc_info=True)
            return False

    def get_active_scheme_guid(self):
        """
        获取当前活动的电源计划的 GUID.

        Returns:
            当前活跃电源计划的 GUID 字符串，或在获取失败时返回 None.
        """
        command = [POWERCARD_COMMAND, GET_ACTIVE_SCHEME_ARG]
        logger.debug(f"Attempting to get active scheme GUID using command: {' '.join(command)}")

        try:
            result = subprocess.run(
                command,
                capture_output=True,
                text=True,
                check=False,
                shell=False,
                creationflags=win32con.CREATE_NO_WINDOW
            )

            logger.debug(f"PowerCfg get active scheme command executed. Return code: {result.returncode}")
            if result.stdout:
                logger.debug(f"PowerCfg get active scheme command stdout: {result.stdout.strip()}")
            if result.stderr:
                # Not typically expected for getactivescheme, but log just in case
                logger.warning(f"PowerCfg get active scheme command stderr: {result.stderr.strip()}")

            # On success (return code 0), parse the stdout
            if result.returncode == 0:
                match = GUID_PARSE_REGEX.search(result.stdout) # re.IGNORECASE is in the compiled regex
                if match:
                    # Extract the captured GUID (group 1) and strip whitespace
                    active_guid = match.group(1).strip()
                    logger.info(f"Successfully retrieved active power scheme GUID: {active_guid}")
                    return active_guid
                else:
                    # Log the output that failed parsing, helpful for debugging regex or unexpected output format
                    logger.error(f"Could not parse GUID from powercfg output: '{result.stdout.strip()}' using regex '{GUID_PARSE_REGEX.pattern}'.")
                    return None
            else:
                logger.error(f"PowerCfg get active scheme command failed with return code {result.returncode} when getting active scheme.")
                return None

        except FileNotFoundError:
            logger.error(f"PowerCfg command '{POWERCARD_COMMAND}' not found. Ensure powercfg is in your system's PATH.", exc_info=True)
            return None
        except Exception as e:
            logger.error(f"An unexpected error occurred while calling powercfg command to get active scheme: {e}", exc_info=True)
            return None

    def get_power_plan_name_from_guid(self, guid: str):
        """
        根据电源计划 GUID 获取对应的名称。

        Args:
            guid: 电源计划的 GUID 字符串。

        Returns:
            对应的电源计划名称，如果在加载的方案中未找到则返回 None。
             注意：此查找是不区分大小写的 GUID 匹配。
        """
        if not guid:
             return None
        # Use lowercase for map lookup (GUIDs are typically case-insensitive when compared)
        return self._guid_to_name_map.get(str(guid).strip().lower())

    def get_available_schemes(self):
        """
        获取加载的所有可用电源方案 (名称 -> GUID 映射).
        返回一个副本。
        """
        return dict(self._name_to_guid_map) # Return a copy

    def get_available_schemes_guid_name_map(self):
        """
        获取加载的所有可用电源方案 (GUID -> 名称 映射).
        GUI 界面可能需要这个映射来显示友好名称。
        返回一个副本。
        """
        return dict(self._guid_to_name_map) # Return a copy
