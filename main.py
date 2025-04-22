import logging
import sys
import os
import threading # Still needed for AppPowerSwitcher, but main thread runs GUI loop
import time # Could be useful for delays, but PumpMessages is blocking
# keyboard is not needed for the basic taskbar icon GUI
# import keyboard # Remove if not used

# datetime is useful for logging and status display
from datetime import datetime

# Important: Adjust sys.path to ensure the 'src' directory is discoverable as a package root.
# This allows absolute imports like 'from src.application...' to work correctly.
# os.path.abspath(__file__) gives the full path to this main.py file.
# os.path.dirname(...) gives the directory containing this file, which is the project root directory.
PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
# Add the project root directory (which contains 'src') to the beginning of sys.path.
# This makes 'src' discoverable as a top-level package.
if PROJECT_ROOT not in sys.path:
    sys.path.insert(0, PROJECT_ROOT)
SRC_DIR = os.path.join(PROJECT_ROOT, "src")
if SRC_DIR not in sys.path:
    sys.path.insert(0, SRC_DIR) # Insert at the beginning for higher priority

# Import pywin32 modules for Taskbar Icon functionality
# These are standard library-level imports for pywin32 installation
try:
    import win32api
    import win32con
    import win32gui
    import winerror # Useful for checking specific Windows errors
except ImportError:
    print("Fatal Error: pywin32 library not found.")
    print("Please install it using: pip install pywin32")
    sys.exit(1)

# Now import necessary modules from our project's src package
try:
    # Import the main application class from the src.application package
    from src.application.power_switcher_app import PowerSwitcherApp
    # Import logging configuration utility from src.utils package
    from src.utils.logging_config import configure_logging
    # Import ConfigManager from src.infrastructure.configuration package
    from src.infrastructure.configuration.config_manager import ConfigManager
    # Need PowerCfgManager to get plan names for GUI display if needed (optional)
    from src.infrastructure.power_management.power_cfg_manager import PowerCfgManager
except ImportError as e:
    print(f"Fatal Error: Could not import core modules from src/. Please ensure 'src' directory structure is correct and imports use 'from src....'.")
    print(f"Details: {e}")
    # Print sys path to help diagnose import issues
    print(f"Current sys.path: {sys.path}")
    sys.exit(1)

# --- Constants for Taskbar Icon ---
# Unique Window Class name for our hidden window messages
TASKBAR_WINDOW_CLASS = "AppPowerSwitcherTrayWindowClass"
# Custom Windows Message ID for taskbar notifications (arbitrary offset)
# Shell_NotifyIcon sends events to the window with this message ID
TASKBAR_NOTIFY_MSG = win32con.WM_USER + 20 # Must be WM_USER + something

# Menu item IDs for the right-click context menu
MENU_EXIT_ID = win32con.WM_USER + 21 # Pick a unique ID for the Exit action

# Taskbar icon ID (can be 0 if only one icon for this window)
TASKBAR_ICON_ID = 0

# Default tooltip text shown when hovering over the icon
TASKBAR_TOOLTIP = "App Power Switcher"

# --- Global reference for the AppPowerSwitcher instance ---
# The TrayIcon callback function needs to access the AppPowerSwitcher instance to stop it.
# We'll store a reference here after it's initialized.
_app_power_switcher_instance = None

# --- Tray Icon Management Class ---
# This class handles creation, management, and destruction of the hidden window and the taskbar icon.
# It runs in the main thread, processing Windows messages.
class TrayIcon:
    """
    管理 Windows 任务栏通知区域图标 (Systray Icon) 及其消息处理。
    """
    def __init__(self, app_switcher: PowerSwitcherApp):
        """
        初始化任务栏图标管理器。

        Args:
            app_switcher: AppPowerSwitcher 应用程序核心实例，用于在退出时停止。
        """
        logger.info("Initializing TrayIcon.")

        # Store reference to the application core for stopping
        if not isinstance(app_switcher, PowerSwitcherApp):
             logger.error("TrayIcon requires an AppPowerSwitcher instance.", exc_info=True)
             raise TypeError("TrayIcon requires an AppPowerSwitcher instance.")

        global _app_power_switcher_instance
        _app_power_switcher_instance = app_switcher # Store global reference for callback

        # 1. Register Window Class (needed to create a window)
        # Use a unique class name to avoid conflicts
        wndClass = win32gui.WNDCLASS()
        wndClass.hInstance = win32api.GetModuleHandle(None)
        wndClass.lpszClassName = TASKBAR_WINDOW_CLASS
        # We need a window procedure (lpfnWndProc) to receive messages
        # We'll use our instance's message map for dynamic message handling
        wndClass.lpfnWndProc = self._message_handler_router # Set the router as the window proc

        # Define handlers for messages the window will receive
        self._message_map = {
            win32con.WM_DESTROY: self.OnDestroy, # Window is being destroyed
            win32con.WM_COMMAND: self.OnCommand, # Menu item selected
            win32con.WM_USER + 20: self.OnTaskbarNotify, # Our custom message ID for taskbar icon events
            win32gui.RegisterWindowMessage("TaskbarCreated"): self.OnTaskbarCreated, # Explorer restarted
            # Add other messages if needed, e.g., WM_QUERYENDSESSION, WM_ENDSESSION for shutdown handling
        }

        try:
            # Try to register the window class
            logger.debug(f"Registering window class: {TASKBAR_WINDOW_CLASS}")
            self._classAtom = win32gui.RegisterClass(wndClass)
            logger.debug("Window class registered successfully.")
        except win32gui.error as e:
            # ERROR_CLASS_ALREADY_EXISTS (1410) is common if app was not stopped cleanly before.
            if e.winerror == winerror.ERROR_CLASS_ALREADY_EXISTS:
                logger.warning(f"Window class '{TASKBAR_WINDOW_CLASS}' already registered. Reusing existing class.")
                # In this case, self._classAtom is not set, but we can proceed.
                # We might need to get the existing class atom if we were strictly correct,
                # but Windows often lets CreateWindow succeed with the name directly.
            else:
                logger.critical(f"Failed to register window class '{TASKBAR_WINDOW_CLASS}': {e}", exc_info=True)
                raise # Re-raise other critical errors

        # 2. Create a hidden Window
        # This window doesn't need to be visible, its purpose is just to receive messages from the taskbar icon.
        style = win32con.WS_OVERLAPPEDWINDOW # Standard window style (can be simple WS_POPUP or 0 too)
        # For a truly invisible window, maybe WS_POPUP | WS_DISABLED
        # But WS_OVERLAPPEDWINDOW is fine, we just don't show it.
        self.hwnd = win32gui.CreateWindow(
            TASKBAR_WINDOW_CLASS,           # Window Class Name
            "App Power Switcher Hidden Window", # Window Title (for identification/debugging)
            style,                          # Window Style
            0, 0,                           # Initial Position (x, y) - use default later
            win32con.CW_USEDEFAULT,         # Initial Size (width) - use default
            win32con.CW_USEDEFAULT,         # Initial Size (height) - use default
            0,                              # Parent Window Handle (0 for desktop)
            0,                              # Menu Handle (0 for none)
            wndClass.hInstance,             # Instance Handle
            None                            # Creation Parameters
        )

        if not self.hwnd:
            # CreateWindow failed
            last_error = win32api.GetLastError()
            logger.critical(f"Failed to create hidden window: HWND is 0. GetLastError: {last_error} ({win32api.FormatMessage(last_error).strip()})", exc_info=True)
            # Unregister class if possible? Or rely on OS cleanup? Proceeding might leak resources.
            # It might be safer to raise an exception here or sys.exit(). Let's raise.
            raise RuntimeError(f"Failed to create hidden window. Last Error: {last_error}")

        logger.info(f"Hidden window created successfully. HWND: {self.hwnd}")

        # Optional: Hide the window explicitly although default sizing might make it small
        # win32gui.ShowWindow(self.hwnd, win32con.SW_HIDE)
        # win32gui.UpdateWindow(self.hwnd) # Update needed sometimes after showing, but hide has no visual update

        # Store a Python object reference with the window handle.
        # This is needed for the C window procedure to find our Python instance.
        # win32gui.SetWindowLong(self.hwnd, win32con.GWL_USERDATA, id(self)) # Use id() to get a unique integer identifier

        # 3. Create Taskbar Icon
        self._create_taskbar_icon()

        logger.info("TrayIcon initialization complete.")

    def _message_handler_router(self, hwnd, msg, wparam, lparam):
        """
        Router function for Windows messages.
        This acts as the main Window Procedure (lpfnWndProc) for our hidden window.
        It looks up the appropriate handler in the _message_map and calls it.
        If no handler is found, it calls the default window procedure.
        """
        # logger.debug(f"Received message: hwnd={hwnd}, msg={msg}, wparam={wparam}, lparam={lparam}") # Very verbose

        # Find the handler for this message in our map
        handler = self._message_map.get(msg)
        if handler:
            # Call the handlers - they should return 0 or 1 for basic messages
            try:
                return handler(hwnd, msg, wparam, lparam)
            except Exception as e:
                 logger.error(f"Error calling message handler for msg {msg}: {e}", exc_info=True)
                 return 0 # Indicate message not processed due to error

        # If no handler is found in our map, pass the message to the default window procedure
        # This is crucial for messages we don't handle, like WM_PAINT, WM_SIZE etc.
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    def _create_taskbar_icon(self):
        """
        Adds the taskbar icon (Notifier Icon Data - NID) to the system tray.
        """
        logger.info("Creating taskbar icon.")
        # Load Icon
        # Can use IDI_APPLICATION as a standard fallback.
        # Or try to load a custom icon file (e.g., app.ico in project root)
        icon_path = os.path.join(PROJECT_ROOT, "app.ico") # Look for app.ico in project root
        hicon = None # Handle to the icon

        if os.path.exists(icon_path):
            try:
                # Load image from file. LR_LOADFROMFILE | LR_DEFAULTSIZE | LR_SHARED
                # LR_SHARED allows reusing the same icon handle across different parts if needed (and is good practice)
                icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE | win32con.LR_SHARED
                hicon = win32gui.LoadImage(
                    win32api.GetModuleHandle(None), # hinst - can be None or module handle
                    icon_path,                      # pszName - file path or resource ID
                    win32con.IMAGE_ICON,            # type - IMAGE_ICON for icon files
                    0, 0,                           # cx, cy - desired width/height (0,0 means use default size specified by flag)
                    icon_flags                      # fuLoad - loading flags
                )
                if hicon:
                     logger.debug(f"Successfully loaded icon from file: {icon_path}")
            except Exception as e:
                logger.warning(f"Failed to load icon from file '{icon_path}': {e}. Using default application icon.", exc_info=True)
                hicon = None # Ensure fallback if file loading failed

        if not hicon:
            # If file loading failed or no file, use a standard Windows icon
            try:
                hicon = win32gui.LoadIcon(0, win32con.IDI_APPLICATION) # Load standard application icon
                logger.debug("Using default application icon.")
            except Exception as e:
                 logger.error(f"Failed to load default application icon: {e}", exc_info=True)
                 # If even default icon fails, proceed without specific icon handle.
                 # Shell_NotifyIcon might use a default blank icon.

        # Define Notifier Icon Data structure (NID)
        # This structure tells Windows about the icon, the associated window, tooltip, etc.
        # See https://learn.microsoft.com/en-us/windows/win32/api/shellapi/ns-shellapi-notifyicondataw
        # We're using NIF_ICON, NIF_MESSAGE, NIF_TIP flags.
        # Need to use the NID structure version that matches the flags we use (NID_V0, V1, V2, V3).
        # NIF_TIP requires NID_V0 or higher. NIF_MESSAGE requires NID_V0. NIF_ICON requires NID_V0.
        # For basic functionality, NID_V0 (or just default structure size calculation) is often enough.
        # Let's use the size of a base NOTIFYICONDATA structure (pre-Vista v3) or the size required by flags.
        # A reliable way is to use the size for the latest version that contains needed flags if possible,
        # but for maximum compatibility with WM_USER+ messages, V0 is implicit by using the older callback style.
        # Let's use win32gui's implicit size calculation based on flags.

        flags = win32gui.NIF_ICON | win32gui.NIF_MESSAGE | win32gui.NIF_TIP
        nid = (
            self.hwnd,          # hWnd: Handle to the window that receives message from the taskbar icon
            TASKBAR_ICON_ID,    # uID: Application-defined identifier for the icon (can be 0)
            flags,              # uFlags: Flags indicating which members are valid
            TASKBAR_NOTIFY_MSG, # uCallbackMessage: The identifier of the application-defined message that Windows sends to the window identified by hWnd
            hicon,              # hIcon: Handle to the icon to display
            TASKBAR_TOOLTIP     # szTip: Tooltip text (max 64 chars including null terminator for V0)
            # For V1/V2/V3 structures, there are more fields and a size field.
            # win32gui.Shell_NotifyIcon seems to handle sizing based on flags context.
        )

        try:
            # Add the icon to the taskbar
            # Use NIM_ADD command
            win32gui.Shell_NotifyIcon(win32gui.NIM_ADD, nid) # Returns TRUE on success, FALSE otherwise.
            logger.info("Taskbar icon created successfully.")
            # Store icon handle if needed for cleanup (though NID handles by hWnd/uID usually)
            # self._hicon = hicon # Keep hicon reference if not LR_SHARED or for explicit DestroyIcon later if needed.
                                 # win32gui may manage hicon internally after NIM_ADD, check docs/examples. LR_SHARED is safer.

        except win32gui.error as e:
            # This can fail if Explorer is not running or taskbar is not ready yet.
            # ERROR_TIMEOUT (1460) might happen
            logger.error(f"Failed to add taskbar icon (Shell_NotifyIcon NIM_ADD): {e}", exc_info=True)
            # The TaskbarCreated message should eventually be sent, triggering OnTaskbarCreated to retry.
            # We continue, but user won't see the icon immediately if this failed.

    # --- Message Handlers ---
    # These methods are called by _message_handler_router when the corresponding Windows message is received.

    def OnDestroy(self, hwnd, msg, wparam, lparam):
        """
        Handler for WM_DESTROY message.
        This message is sent when the window is being destroyed.
        Clean up the taskbar icon and post a WM_QUIT message to exit the message loop.
        """
        logger.info(f"Received WM_DESTROY message for HWND {hwnd}. Cleaning up.")

        # 1. Delete the taskbar icon
        # Use NIM_DELETE command
        nid = (self.hwnd, TASKBAR_ICON_ID) # Need hWnd and uID to specify which icon to delete
        try:
            win32gui.Shell_NotifyIcon(win32gui.NIM_DELETE, nid)
            logger.info("Taskbar icon deleted successfully.")
        except win32gui.error as e:
            logger.warning(f"Failed to delete taskbar icon (Shell_NotifyIcon NIM_DELETE): {e}", exc_info=True)
            # Continue cleanup despite this.

        # 2. Post WM_QUIT message
        # This message signals the message loop (win32gui.PumpMessages) to exit.
        logger.debug("Posting WM_QUIT message to exit message loop.")
        win32gui.PostQuitMessage(0) # The parameter is the exit code (0 for success)
        logger.info("WM_QUIT message posted.")

        # Note: The _app_power_switcher_instance stop() is called *before* DestroyWindow
        # from the menu handler. So cleanup should occur before WM_DESTROY fires.

        return 0 # Return 0 for WM_DESTROY message

    def OnCommand(self, hwnd, msg, wparam, lparam):
        """
        Handler for WM_COMMAND message.
        This message is sent when a menu item is selected, a control sends a notification, etc.
        We are interested in menu item selections.
        """
        # win32api.LOWORD(wparam) is the control ID or menu item ID
        command_id = win32api.LOWORD(wparam)
        # win32api.HIWORD(wparam) is the notification code (0 for menu)
        # lparam is the handle to the control (0 for menu)

        logger.debug(f"Received WM_COMMAND message. Command ID: {command_id}")

        if command_id == MENU_EXIT_ID:
            # The "退出" menu item was selected
            logger.info(f"'退出' menu item (ID {MENU_EXIT_ID}) selected. Initiating application stop.")
            # 1. Stop the core AppPowerSwitcher logic (background threads, hooks, etc.)
            try:
                # Use the global reference to the AppPowerSwitcher instance
                if _app_power_switcher_instance and _app_power_switcher_instance._running.is_set():
                    logger.info("Calling AppPowerSwitcher stop() from menu handler.")
                    _app_power_switcher_instance.stop()
                    logger.info("AppPowerSwitcher stop() method finished.")
                elif _app_power_switcher_instance:
                     logger.warning("AppPowerSwitcher instance was already stopped or not running when 'Exit' was selected.")
                else:
                    logger.error("AppPowerSwitcher instance reference is not available when 'Exit' was selected. Cannot stop cleanly.")

            except Exception as e:
                 logger.error(f"Error during AppPowerSwitcher stop() from menu handler: {e}", exc_info=True)
                 # Continue with window destruction attempt even if stop failed poorly

            # 2. Destroy the hidden window
            # This will cause WM_DESTROY to be sent, which performs the final GUI cleanup (deleting icon, posting WM_QUIT).
            logger.info("Destroying hidden window to exit message loop.")
            win32gui.DestroyWindow(self.hwnd)
            # WM_DESTROY handler will take over from here to complete exit.

        else:
            # Handle other potential WM_COMMAND messages if any controls were added to the window (unlikely for hidden window)
            logger.debug(f"Unhandled WM_COMMAND ID: {command_id}")

        return 0 # Return 0 for WM_COMMAND message

    def OnTaskbarNotify(self, hwnd, msg, wparam, lparam):
        """
        Handler for our custom TASKBAR_NOTIFY_MSG (WM_USER + 20).
        This message carries information about actions on the taskbar icon.
        lparam contains the mouse message (WM_LBUTTONUP, WM_RBUTTONUP, WM_LBUTTONDBLCLK, etc.).
        """
        # logger.debug(f"Received TaskbarNotify ({TASKBAR_NOTIFY_MSG}). lparam: {lparam} (Mouse msg)") # Very verbose

        if lparam == win32con.WM_RBUTTONUP:
            # Right-click on the icon
            logger.debug("Taskbar icon Right-clicked. Showing context menu.")
            # 1. Create the popup menu
            menu = win32gui.CreatePopupMenu()

            # 2. Append menu items
            # win32gui.AppendMenu(hMenu, uFlags, uIDNewItem, lpNewItem)
            # uFlags: MF_STRING for text item, etc.
            # uIDNewItem: Command ID that will be sent in WM_COMMAND when selected
            # lpNewItem: The text for the menu item
            win32gui.AppendMenu(menu, win32con.MF_STRING, MENU_EXIT_ID, "退出") # Only the "Exit" option

            # Optional: Add a separator or other items if needed later
            # win32gui.AppendMenu(menu, win32con.MF_SEPARATOR, 0, None)
            # win32gui.AppendMenu(menu, win32con.MF_STRING, SOME_OTHER_ID, "Another Option")

            # 3. Get mouse position to display menu there
            pos = win32gui.GetCursorPos()

            # 4. Track the popup menu
            # Set the foreground window to receive menu messages correctly.
            win32gui.SetForegroundWindow(self.hwnd)
            # TrackPopupMenu(hMenu, uFlags, x, y, res, hwnd, lprc)
            # uFlags: TPM_LEFTALIGN, TPM_RIGHTBUTTON etc.
            win32gui.TrackPopupMenu(
                menu,                      # hMenu: Handle to the popup menu
                win32con.TPM_LEFTALIGN |   # uFlags: Alignment flags
                win32con.TPM_RIGHTBUTTON, # Use right button to track (required by some styles)
                pos[0], pos[1],            # x, y: Screen coordinates for the top-left of the menu
                0,                         # nReserved: Must be 0
                self.hwnd,                 # hWnd: Window to receive menu messages (our hidden window)
                None                       # lprc: Optional rectangle
            )
            # Post a dummy message to the window. This is a workaround needed in some cases
            # after using TrackPopupMenu to ensure the window receives the WM_COMMAND message correctly.
            win32gui.PostMessage(self.hwnd, win32con.WM_NULL, 0, 0)

            # 5. Destroy the menu after use (important for resource cleanup)
            win32gui.DestroyMenu(menu)
            logger.debug("Popup menu displayed.")

        elif lparam == win32con.WM_LBUTTONUP:
             # Left-click on the icon
             logger.debug("Taskbar icon Left-clicked.")
             # You could add logic here, e.g., show the main GUI configuration window if it existed.
             # For now, let's just log.
             pass # Or maybe show a balloon tip status?

        elif lparam == win32con.WM_LBUTTONDBLCLK:
             # Double-left-click on the icon
             logger.debug("Taskbar icon Double-clicked (Left).")
             # Default action could be to show config window or exit.
             # Let's make double-click also initiate exit for convenience.
             # This calls the same logic as the "退出" menu item.
             self.OnCommand(hwnd, win32con.WM_COMMAND, MENU_EXIT_ID, 0)

        # Return 1 to indicate message was handled
        return 1

    def OnTaskbarCreated(self, hwnd, msg, wparam, lparam):
        """
        Handler for the "TaskbarCreated" registered message.
        This message is broadcast when the taskbar has been recreated (e.g., Explorer crashes and restarts).
        We need to recreate our taskbar icon in this scenario.
        """
        logger.warning("TaskbarCreated message received. Recreating taskbar icon.")
        # Delete the old icon data (it may be invalid now) and recreate the icon.
        # Shell_NotifyIcon(NIM_DELETE) is not strictly necessary if the old icon is already gone,
        # but might be safe. Let's just call create directly, it should handle internally.
        self._create_taskbar_icon()
        return 0

# --- Global Window Message Hook for Python Instance Routing ---
# Since win32gui.WNDCLASS.lpfnWndProc needs a C-compatible function pointer,
# it's difficult to directly point it to a bound method like self.OnMessage.
# A common technique is to map the C function pointer to a global/static function
# or a factory function that can find the Python instance associated with the HWND.
# The demo code did this by setting the message map directly on the wc object's lpfnWndProc.
# This simple approach works for basic message handling where the handler signature matches DefWindowProc.
# Our _message_handler_router method has the correct signature (hwnd, msg, wparam, lparam).
# So, wc.lpfnWndProc = self._message_handler_router in __init__ is the correct way to bind it for *that specific window*.
# This does NOT require a global hook or complex instance lookup if doing it per window creation.

# The trick in the demo was that the WNDCLASS registration defined the lpfnWndProc *before* the window was created.
# Win32gui allows setting lpfnWndProc to a Python function object directly.

# Let's verify this works by leaving the `wc.lpfnWndProc = self._message_handler_router` line as is.

# --- Main Application Entry Point ---

# Initial configuration of logging is crucial and should happen here first thing.
# The AppPowerSwitcher will load config and might implicitly try to configure logging again,
# but calling configure_logging here first ensures the root logger is set up correctly
# with handlers and basic level.
# A temporary ConfigManager instance is used *only* to get the desired log level early.

print("AppPowerSwitcher main script started.")
print("Loading initial configuration for logging level...")
# Create a temporary ConfigManager instance just to read the config file for log level
# ConfigManager uses its internal logic to find the config file path relative to the determined PROJECT_ROOT
temp_config_manager_for_log = ConfigManager()
# Load config. This will also handle creating the default file if it doesn't exist.
load_success_for_log = temp_config_manager_for_log.load_config()

# Determine the desired log level string from the loaded config (or use a fallback like "INFO" if loading failed)
log_level_str = temp_config_manager_for_log.get_log_level() if load_success_for_log else "INFO"
# Convert the log level string (e.g., "INFO") to the corresponding logging module constant (e.g., logging.INFO)
log_level = getattr(logging, log_level_str.upper(), logging.INFO)

# Configure the root logging system based on the determined level. All loggers inherit from this.
print(f"Configuring shared logging system with level: {log_level_str}")
configure_logging(log_level=log_level)

# Now that logging is fully configured, retrieve the logger for the main script itself.
logger = logging.getLogger(__name__)
logger.info("Logging system initialized successfully via main.py.")

def main():
    """
    主函数：应用程序的入口。
    加载配置，启动核心应用逻辑 (后台线程)，创建任务栏图标GUI (主线程)，运行消息循环。
    """
    logger.info("-" * 40)
    logger.info("Starting AppPowerSwitcher application...")
    logger.info("-" * 40)

    # --- Permission Requirement ---
    # Power plan switching requires Administrator privileges.
    # The GUI (creating window, taskbar icon) generally does NOT require admin.
    # If the app starts without admin, the GUI will show, but power switching will fail repeatedly (logged).
    # It's important to inform the user.
    print("\n### IMPORTANT: Power plan switching requires Administrator privileges! ###")
    print("### Please ensure you run this script as Administrator if switching fails. ###\n")
    logger.warning("Application starting. Power plan switching requires Administrator privileges.")

    app_switcher = None # Initialize main app core instance
    tray_icon = None    # Initialize tray icon instance

    try:
        # 1. Instantiate and Start the core application logic (runs in background threads)
        # This loads full config, starts power manager, event listener, and processing thread.
        app_switcher = PowerSwitcherApp()
        logger.info("AppPowerSwitcher instance created.")

        app_start_success = app_switcher.start()
        if not app_start_success:
             logger.critical("AppPowerSwitcher failed to start successfully. Exiting application.")
             # AppPowerSwitcher.start() logs internal reasons for failure.
             # Its stop() is already called internally upon critical startup failure.
             return # Exit main function

        logger.info("AppPowerSwitcher core started successfully in background threads.")

        # 2. Create and manage the Taskbar Icon GUI (runs in the main thread)
        # Pass the AppPowerSwitcher instance to the TrayIcon manager
        tray_icon = TrayIcon(app_switcher)
        logger.info("TrayIcon initialized.")

        # 3. Run the Windows Message Loop.
        # This call blocks the main thread and makes it responsible for processing messages
        # sent to the hidden window (which come from user clicks on icon, or system events).
        # Our event handlers will receive messages here and call the appropriate logic (like app_switcher.stop()).
        logger.info("Entering Windows message loop (win32gui.PumpMessages).")
        # PumpMessages processes messages until win32gui.PostQuitMessage is called.
        win32gui.PumpMessages()
        logger.info("Windows message loop exited.")

    except KeyboardInterrupt:
        # Handle Ctrl+C gracefully. This exception occurs in the main thread before PumpMessages starts or while it's running.
        logger.info("Keyboard interrupt (Ctrl+C) received in main thread. Application stopping.")
        # If PumpMessages was running, it was likely interrupted.
        # We need to ensure cleanup happens. The finally block will handle calling app_switcher.stop().
        # We don't need to explicitly post WM_QUIT here if PumpMessages was running, as the Ctrl+C likely broke it out.

    except Exception as e:
         # Catch any unexpected critical errors during setup or before PumpMessages starts or during PumpMessages execution itself if it doesn't catch it.
         logger.critical(f"A critical error occurred during application execution: {e}", exc_info=True)
         # Ensure cleanup happens. The finally block will handle calling app_switcher.stop().

    finally:
        # Ensure app is stopped cleanly when main function exits, regardless of how.
        # This block is executed upon normal loop exit (WM_QUIT), break from loop (thread death), or exceptions.
        # Check if app_switcher instance was created and if it was still logically running (its _running flag is set).
        # If the "退出" menu item was used, app_switcher.stop() would have been called, clearing _running.
        # If we exit due to Ctrl+C or unexpected thread death, _running might still be set.
        if app_switcher and app_switcher._running.is_set():
             logger.info("Main function cleanup: App was still marked as running. Calling app_switcher.stop() gracefully.")
             app_switcher.stop()
        elif app_switcher and not app_switcher._running.is_set():
             # This case happens if app_switcher.stop() was already called (e.g., by the "Exit" menu handler)
             logger.info("Main function cleanup: App was already stopped gracefully before final cleanup.")
        else:
             # This case happens if an exception occurred *before* app_switcher instance was fully initialized or started.
             logger.info("Main function cleanup: Application instance was not fully initialized or in an unknown state.")

        # Note: TrayIcon cleanup (Shell_NotifyIcon NIM_DELETE) is handled by TrayIcon.OnDestroy,
        # which is triggered by win32gui.DestroyWindow(self.hwnd) called in OnCommand.
        # No explicit TrayIcon cleanup call is needed in this finally block.

    logger.info("-" * 40)
    logger.info("AppPowerSwitcher application finished.")
    logger.info("-" * 40)

if __name__ == "__main__":
    # This is the actual entry point when the script is executed.
    main()