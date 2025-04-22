import logging
import threading
import queue
import time
import sys
import os # Needed for potential sys.path manipulation in fallback logging
# datetime is needed for status print (especially in get_current_state_info)
from datetime import datetime

# Import modules from the infrastructure layer
# Assuming src is in the sys.path when this module is imported (handled by main.py)
try:
    from src.infrastructure.configuration.config_manager import ConfigManager
    from src.infrastructure.power_management.power_cfg_manager import PowerCfgManager
    from src.infrastructure.windows.event_listener import EventListener
    # Configure logging utility is primarily called by main.py, but mentioned here
    # from utils.logging_config import configure_logging # Not directly called in App, but depends on it being called elsewhere
except ImportError as e:
    print(f"Fatal Error: Could not import infrastructure or utils modules: {e}")
    print("Please ensure you are running from the project root directory 'AppPowerSwitcher/'")
    print("e.g., 'python main.py'") # Corrected example command
    sys.exit(1)

# Get logger for this module
# This logger will pick up the logging configuration set up by main.py
logger = logging.getLogger(__name__)

# --- PowerSwitcherApp Class ---

class PowerSwitcherApp:
    """
    应用程序核心类，负责协调配置加载、事件监听和电源计划管理。
    它作为前台界面和后台操作之间的桥梁。
    使用电源计划的 GUID 作为主要标识符。
    """
    def __init__(self):
        """
        初始化应用程序核心组件。
        实例化配置管理器、电源管理器和事件监听器。
        """
        logger.info("Initializing PowerSwitcherApp.")

        # Infrastructure components
        # ConfigManager is instantiated early, it will load the config file (or create default)
        # It uses its internal logic to find the config file path relative to the project root
        self._config_manager = ConfigManager()

        # PowerCfgManager needs to load system power schemes (GUIDs and Names)
        # Loading happens in PowerCfgManager's __init__ when it's instantiated
        self._power_manager = PowerCfgManager()

        # Queue for communication between EventListener (producer) and the Processing thread (consumer)
        # EventListener puts foreground process names (strings) into this queue.
        # Processing thread gets process names from this queue.
        # Limit queue size to prevent excessive memory growth if processing is slow.
        self._event_queue = queue.Queue(maxsize=100) # Use a larger queue size if needed for buffering

        # EventListener receives window events and puts process names into the queue
        # It needs the queue instance to send data.
        self._event_listener = EventListener(self._event_queue)

        # Thread to process items (process names) from the queue
        self._processing_thread = None
        # Event flag to signal the processing thread to continue running (set by start, cleared by stop)
        self._running = threading.Event()

        # State variables to track last processed info to avoid unnecessary actions
        self._last_processed_process = None # Last process name received from event listener
        # _last_applied_power_plan_identifier will store the GUID string
        self._last_applied_power_plan_identifier = None
        # _last_known_active_guid stores the validated GUID of the plan that was actually active
        # after the last check/switch attempt.
        self._last_known_active_guid = None

        logger.info("PowerSwitcherApp initialized.")

    def start(self):
        """
        启动应用程序核心逻辑。
        加载配置，加载电源方案 (已在 __init__ 中完成)，启动事件监听和处理线程。
        """
        logger.info("Starting PowerSwitcherApp.")

        # 1. Load configuration
        # The ConfigManager was already instantiated in __init__, loading config at that point.
        # We can optionally reload or just confirm it was loaded.
        # Let's add a check here to see if config load was successful.
        # If config was created because it didn't exist, load_config already ran and created it.
        # If config file existed, load_config ran and loaded it.
        # If critical failure during config loading, load_config would return False.
        # Let's ensure config is loaded again here just before starting threads,
        # in case settings changed before app.start() was called (e.g., via GUI setting up config first).
        load_success = self._config_manager.load_config() # Reload config
        if not load_success:
            logger.error("Failed to (re)load configuration during app startup. Application cannot start reliably.")
            # Decide on fatal vs continue. For now, continuing with defaults as loaded in __init__.
            # A critical app might exit here.
            # return False # Or raise Exception()

        # 2. Configure logging based on loaded config (should already be done by main.py, but check is safe)
        # Our main.py does this now, so this check should find handlers.
        root_logger = logging.getLogger()
        if not root_logger.handlers:
             logger.warning("Logging was NOT configured by main.py before App init. Using settings from config loaded here.")
             # This case should not happen if main.py is structured correctly.
             # If it does, try to configure logging based on the now loaded config.
             log_level_str = self._config_manager.get_log_level()
             log_level = getattr(logging, log_level_str.upper(), logging.INFO)
             # configure_logging() needs access to file paths, etc. which it gets internally
             # If main.py couldn't configure, configure_logging here might also struggle if permissions/paths are bad.
             # Just log a warning and trust the initial logging setup in main.py.
             # configure_logging(log_level=log_level) # Avoid calling configure_logging here if main.py already did

        # 3. Start PowerCfgManager (schemes loaded in __init__)
        # Check if any schemes were loaded successfully by the manager's __init__
        available_schemes = self._power_manager.get_available_schemes_guid_name_map() # Use GUID->Name map for display in log
        if not available_schemes:
             logger.warning("No power schemes were loaded by PowerCfgManager. Power switching functionality may be limited or non-functional.")
        else:
             # Log loaded schemes for verification
             logger.info(f"Available Power Schemes loaded by PowerManager: {available_schemes}")

        # 4. Start EventListener (will run in a separate thread and set the Windows Hook)
        try:
            self._event_listener.start()
            logger.info("EventListener thread requested to start.")
        except Exception as e:
             logger.critical(f"Failed to start EventListener thread: {e}", exc_info=True)
             # If listener fails, the core loop won't receive events. Application cannot function.
             self.stop() # Attempt to stop everything else cleanly (will mostly just log warnings about things not running)
             raise RuntimeError("Failed to start EventListener, exiting.") from e # Re-raise as critical error

        # 5. Start the processing thread (will retrieve events from queue)
        self._running.set() # Set the running flag before starting the thread
        self._processing_thread = threading.Thread(target=self._process_queue, name="ProcessingThread")
        # Processing thread should probably NOT be daemon if we want clean shutdown.
        # If daemon=False, main thread must join() it on exit. Our stop() method tries to join.
        # If daemon=True, it might exit abruptly without finishing current queue item or cleanup.
        # Let's keep it daemon=False for better control during stop().
        self._processing_thread.daemon = False
        self._processing_thread.start()
        logger.info("ProcessingThread requested to start.")

        logger.info("PowerSwitcherApp start sequence finished. Core threads are expected to be running.")
        return True # Indicate startup was successful

    def stop(self):
        """
        停止应用程序核心逻辑。
        向处理线程发送停止信号，停止事件监听，等待线程退出。
        """
        # Check if the running flag is set. If not, we are already stopping or stopped.
        if not self._running.is_set():
            logger.warning("PowerSwitcherApp is already stopping or is stopped.")
            # Even if the flag is clear, attempt cleanups for robustness.
            self._cleanup_threads_and_hooks()
            return

        logger.info("Stopping PowerSwitcherApp.")

        # 1. Signal the processing thread to stop. Clear the running flag.
        self._running.clear() # Clear the running flag (main loop condition in _process_queue)

        # 2. Unblock the processing thread's queue.get() by putting a sentinel value (None).
        # Use put_nowait to ensure this doesn't block the stop() method itself.
        try:
            if self._event_queue: # Check if queue was initialized
                 self._event_queue.put_nowait(None)
                 logger.debug("Put sentinel value (None) into the event queue to unblock processing thread.")
            else:
                 logger.warning("Event queue not initialized during stop sequence.")
        except queue.Full:
            logger.warning("Event queue is full during stop sequence, could not put sentinel. Processing thread might take longer to stop.")
        except Exception as e:
            logger.error(f"Error putting sentinel into queue during stop: {e}", exc_info=True)

        # 3. Perform cleanup sequence (Unhook, join threads)
        # Consolodate cleanup logic into a separate method
        self._cleanup_threads_and_hooks()

        logger.info("PowerSwitcherApp stopped.")

    def _cleanup_threads_and_hooks(self):
        """
        Helper method to stop child threads and unhook Windows events.
        Called by stop() or if startling fails.
        """
        logger.info("Initiating PowerSwitcherApp cleanup sequence.")

        # 1. Stop EventListener (this sends WM_QUIT to its message loop thread and unhooks)
        try:
            if self._event_listener: # Check if event listener was initialized
                 self._event_listener.stop()
                 logger.info("EventListener stop method called.")
                 # stop() method handles joining its internal thread.
            else:
                 logger.warning("EventListener not initialized during cleanup sequence.")
        except Exception as e:
             logger.error(f"Failed to stop EventListener cleanly: {e}", exc_info=True)
             # Continue cleanup of other parts

        # 2. Wait (join) for the processing thread to finish
        # The signal (_running.clear() + sentinel in queue) should cause it to exit its loop.
        if self._processing_thread and self._processing_thread.is_alive():
            logger.debug("Waiting for ProcessingThread to join during cleanup...")
            # Join with a timeout to prevent main thread from hanging indefinitely
            timeout_sec = 5.0
            self._processing_thread.join(timeout=timeout_sec)

            if self._processing_thread.is_alive():
                 logger.error(f"ProcessingThread did not stop cleanly within {timeout_sec} seconds timeout.")
            else:
                 logger.info("ProcessingThread joined successfully.")
        elif self._processing_thread:
             logger.debug("ProcessingThread was not alive or not initialized during cleanup.")
        else:
             logger.debug("ProcessingThread was not initialized.")

        logger.info("PowerSwitcherApp cleanup sequence finished.")

    def _process_queue(self):
        """
        Worker function executed by the processing thread.
        Retrieves process names from the queue and triggers power plan switching logic
        based on configured GUIDs.
        """
        # Ensure the logger is available in this thread's context (usually works fine with module-level logger)
        # log_thread = logging.getLogger(__name__) # Alternative way to get logger in thread
        logger.info("ProcessingThread started queue processing loop (using GUIDs).")

        # Main processing loop runs while the application's _running flag is set
        # _running.is_set() is controlled by app.start() and app.stop()
        while self._running.is_set():
            try:
                # Get an item from the queue with a short timeout.
                # The timeout allows the loop to check the self._running signal periodically
                # even if no items are in the queue. This is crucial for responsiveness during stop.
                # Setting block=True is default, timeout makes it return queue.Empty on timeout.
                process_name = self._event_queue.get(block=True, timeout=0.5)

                # Check if the retrieved item is the sentinel value signaling stop
                if process_name is None:
                    logger.debug("ProcessingThread received stop sentinel (None). Exiting loop.")
                    break # Exit the while loop

                # --- Core Business Logic ---
                # We received a valid foreground process name string.
                logger.debug(f"Processing received process name from queue: {process_name}")

                # Avoid processing the same process name repeatedly UNLESS the configuration might have changed.
                # With hot-reloading config (future feature), we'd need to re-check the config even for the same process.
                # For now, simple check: if the process name is the same as the last one processed, skip.
                if process_name == self._last_processed_process:
                     # Log info level if you want to see every time this optimization kicks in
                     logger.debug(f"Processed same process again: '{process_name}'. Skipping power plan check.")
                     continue # Skip lookup and switch logic

                self._last_processed_process = process_name # Update last processed process tracker

                # 1. Look up the desired power plan GUID for this process in the configuration
                # ConfigManager's method handles case-insensitive lookup and returns the GUID string or None.
                target_power_plan_guid = self._config_manager.get_power_plan_for_process(process_name)

                if target_power_plan_guid is None:
                    # If the process is not in the map, get the default power plan GUID from config.
                    logger.debug(f"Process '{process_name}' not found in config map. Retrieving default power plan GUID.")
                    target_power_plan_guid = self._config_manager.get_default_power_plan()
                    # Check if the default GUID is also empty or None (unlikely with ConfigManager fallbacks, but defensive)
                    if not target_power_plan_guid:
                        logger.warning(f"Default power plan GUID for process '{process_name}' is empty or None from config. Cannot apply default plan.")
                         # Log the issue and process the next event from the queue.
                        continue # Skip switch attempt

                 # Ensure resolved target_power_plan_guid is a non-empty string after cleaning whitespace
                target_power_plan_guid = str(target_power_plan_guid).strip()
                if not target_power_plan_guid:
                     logger.warning(f"Resolved target power plan GUID for process '{process_name}' is empty or whitespace only. Cannot switch.")
                     continue

                logger.debug(f"Target power plan GUID for '{process_name}': {target_power_plan_guid}")

                # 2. Get the current active power plan GUID from the system
                current_active_guid = self._power_manager.get_active_scheme_guid()

                if not current_active_guid:
                    logger.error("Failed to get current active power scheme GUID from system. Cannot determine if switch is needed.")
                    # Processed event, but couldn't get current state. Continue.
                    continue # Skip switch attempt

                # 3. Compare current active GUID with the target GUID (case-insensitive comparison)
                # Both are expected to be valid GUID strings at this point.
                # Use .lower() for robust comparison as GUIDs are case-insensitive.
                if current_active_guid.lower() == target_power_plan_guid.lower():
                     # Current plan is already the desired target plan or its GUID equivalent.
                     current_active_name = self._power_manager.get_power_plan_name_from_guid(current_active_guid) or "Unknown"
                     target_name = self._power_manager.get_power_plan_name_from_guid(target_power_plan_guid) or "Unknown" # Try to get friendly name for log
                     logger.debug(f"Current active plan (GUID: {current_active_guid}, Name: '{current_active_name}') is already the target plan for '{process_name}' (Target GUID: {target_power_plan_guid}, Name: '{target_name}'). No switch needed.")

                     # Although no switch occurred, update our internal last applied/known state
                     # if the effective target plan (identified by its GUID) is new.
                     if self._last_known_active_guid != target_power_plan_guid.lower():
                         # This case means the active plan was already the desired one, but it might be a different plan than the *last one we explicitly switched to*.
                         # Update trackers to reflect current discovered state.
                         logger.debug(f"Updating _last_applied_power_plan_identifier and _last_known_active_guid based on current active plan match for process '{process_name}'.")
                         self._last_applied_power_plan_identifier = target_power_plan_guid # Store the target GUID
                         self._last_known_active_guid = target_power_plan_guid.lower() # Store lowercase target GUID for comparison
                     continue # Skip switching

                # 4. If current active plan is different from the target, attempt to switch.
                current_active_name = self._power_manager.get_power_plan_name_from_guid(current_active_guid) or "Unknown"
                target_name = self._power_manager.get_power_plan_name_from_guid(target_power_plan_guid) or "Unknown"
                logger.info(f"Current active plan (GUID: {current_active_guid}, Name: '{current_active_name}') is different from target for '{process_name}' (GUID: {target_power_plan_guid}, Name: '{target_name}'). Attempting to switch.")

                # Call the power manager's switch method, passing the target GUID.
                # PowerCfgManager expects a GUID string and calls powercfg /setactive.
                switch_success = self._power_manager.switch_power_plan(target_power_plan_guid)

                if switch_success:
                    logger.info(f"Power plan switch requested successfully for '{process_name}' to GUID '{target_power_plan_guid}'.")
                    # Update internal state trackers after a successful switch request.
                    # The actual plan might take a moment to apply in the OS, but we requested it successfully.
                    self._last_applied_power_plan_identifier = target_power_plan_guid # Store the target GUID
                    # Update last known GUID to the target we aimed for after successful request
                    self._last_known_active_guid = target_power_plan_guid.lower()
                else:
                    logger.error(f"Failed to switch power plan for '{process_name}' to GUID '{target_power_plan_guid}'. Check power_manager logs for details. (Likely permissions or invalid GUID).")
                    # Do NOT update self._last_applied_power_plan_identifier or _last_known_active_guid, as the switch failed.

                # Finished processing this event. The loop continues to wait for the next event.

            except queue.Empty:
                # This exception is expected periodically due to the timeout in get().
                # It allows the thread to check the while condition (_running.is_set()).
                pass
            except Exception as e:
                # Catch any unexpected errors within the processing loop itself
                logger.error(f"An unexpected error occurred in the ProcessingThread queue loop while processing event for process '{process_name if 'process_name' in locals() else 'N/A'}': {e}", exc_info=True)
                # Log the error and continue the loop to process the next event.
                # A severe, recurring error might indicate a fundamental issue that needs
                # more robust error handling or a mechanism to stop the thread.

        # The while loop condition "_running.is_set()" became False, or hit a break condition.
        logger.info("ProcessingThread queue processing loop finished.")

    # --- Methods might be called by GUI layer ---
    # These methods provide interfaces for the GUI to interact with the app's state and functionality.

    def get_config_manager(self):
        """
        Provides access to the ConfigManager instance.
        Useful for the GUI to load, display, and save configuration.
        The ConfigManager holds access to config data (e.g., process-GUID map, default GUID).
        """
        return self._config_manager

    def get_power_manager(self):
        """
        Provides access to the PowerCfgManager instance.
        Useful for the GUI to potentially list available schemes (GUIDs and Names)
        or check current plan for display.
        """
        return self._power_manager

    def get_current_state_info(self):
        """
        Provides current operational state information for display in GUI.
        Includes current active power plan (GUID and Name), last processed process, etc.
        """
        try:
            # Get current active plan GUID from PowerManager
            current_guid = self._power_manager.get_active_scheme_guid()
            # Try to resolve GUID to friendly name using the map loaded by PowerCfgManager
            current_name = self._power_manager.get_power_plan_name_from_guid(current_guid) if current_guid else "Unknown"

            # Get name for the last applied identifier (which is now expected to be a GUID)
            last_applied_guid = self._last_applied_power_plan_identifier
            last_applied_name = self._power_manager.get_power_plan_name_from_guid(last_applied_guid) if last_applied_guid else "None"

            return {
                "is_running": self._running.is_set(),
                "processing_thread_alive": self._processing_thread is not None and self._processing_thread.is_alive(),
                "event_listener_thread_alive": self._event_listener is not None and hasattr(self._event_listener, '_listener_thread') and self._event_listener._listener_thread is not None and self._event_listener._listener_thread.is_alive(),
                "last_processed_process": self._last_processed_process,
                "last_applied_power_plan": f"'{last_applied_name}' ({last_applied_guid})" if last_applied_guid else "None", # Display both name and GUID
                "current_active_power_plan": f"'{current_name}' ({current_guid})" if current_guid else "Unknown (N/A)", # Display both name and GUID
                "queue_size": self._event_queue.qsize() if self._event_queue else "N/A",
            }
        except Exception as e:
             # Log the error but try to return some basic info even if some parts fail
             logger.error(f"Error getting current state information: {e}", exc_info=True)
             # Return a dictionary with error indicator and whatever basic state we can access
             return {
                 "error": "Error retrieving status: " + str(e),
                 "is_running": self._running.is_set() if hasattr(self, '_running') else False,
                 "processing_thread_alive": self._processing_thread is not None and self._processing_thread.is_alive() if hasattr(self, '_processing_thread') else False,
                 "event_listener_thread_alive": self._event_listener is not None and hasattr(self._event_listener, '_listener_thread') and self._event_listener._listener_thread is not None and self._event_listener._listener_thread.is_alive() if hasattr(self, '_event_listener') else False,
                 "last_processed_process": self._last_processed_process if hasattr(self, '_last_processed_process') else None,
                 "last_applied_power_plan": "N/A (Status Error)",
                 "current_active_power_plan": "N/A (Status Error)",
                 "queue_size": self._event_queue.qsize() if hasattr(self, '_event_queue') and self._event_queue else "N/A",
             }
