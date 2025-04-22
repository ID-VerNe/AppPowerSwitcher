import logging
import os
import sys
import ctypes
import ctypes.wintypes
import win32con # We can still use win32con for constants like PROCESS_QUERY_INFORMATION

# Get logger for this module
logger = logging.getLogger(__name__)

# --- Define ctypes functions and types needed ---

# Load required Windows DLLs
user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# Define necessary Windows data types for ctypes
DWORD = ctypes.wintypes.DWORD
HWND = ctypes.wintypes.HWND
HANDLE = ctypes.wintypes.HANDLE
LPWSTR = ctypes.wintypes.LPWSTR # Pointer to a wide string
LPDWORD = ctypes.wintypes.LPDWORD

# Define function signatures for ctypes
# BOOL GetWindowThreadProcessId(HWND hWnd, LPDWORD lpdwProcessId);
user32.GetWindowThreadProcessId.argtypes = [HWND, LPDWORD] # LPWSTR because LPDWORD is essentially output pointer
user32.GetWindowThreadProcessId.restype = DWORD # Returns thread ID

# HANDLE OpenProcess(DWORD dwDesiredAccess, BOOL bInheritHandle, DWORD dwProcessId);
kernel32.OpenProcess.argtypes = [DWORD, ctypes.wintypes.BOOL, DWORD]
kernel32.OpenProcess.restype = HANDLE

# BOOL CloseHandle(HANDLE hObject);
kernel32.CloseHandle.argtypes = [HANDLE]
kernel32.CloseHandle.restype = ctypes.wintypes.BOOL

# DWORD QueryFullProcessImageName(HANDLE hProcess, DWORD dwFlags, LPWSTR lpExeName, PDWORD lpdwSize);
# Using W version for Unicode
kernel32.QueryFullProcessImageNameW.argtypes = [HANDLE, DWORD, LPWSTR, LPDWORD]
kernel32.QueryFullProcessImageNameW.restype = ctypes.wintypes.BOOL # Success/Failure

# Define flags (can still use win32con values)
PROCESS_QUERY_INFORMATION = win32con.PROCESS_QUERY_INFORMATION # 0x0400
PROCESS_VM_READ = win32con.PROCESS_VM_READ                     # 0x0010 - sometimes needed, though QUERY_INFORMATION usually works

# Combine ACCESS_MASK for OpenProcess if needed (though QUERY_INFORMATION should be enough)
PROCESS_ACCESS = PROCESS_QUERY_INFORMATION # | PROCESS_VM_READ # Use | if also need PROCESS_VM_READ

def get_process_name_from_hwnd(hwnd):
    """
    根据 Windows 窗口句柄 (HWND) 获取所属进程的可执行文件名称 (.exe).
    使用 ctypes 调用 Windows API.

    Args:
        hwnd: Windows 窗口句柄 (integer).

    Returns:
        如果成功获取到进程名称，返回进程文件名称 (例如 "notepad.exe")。
        如果获取失败 (例如句柄无效、权限不足等)，返回 None。
    """
    logger.debug(f"Entering get_process_name_from_hwnd (ctypes): hwnd={hwnd}")

    # Check if hwnd is potentially valid (basic check)
    # Note: user32.IsWindow is also available via ctypes if needed, but relying on API calls to fail is also common.
    if not hwnd or hwnd == 0:
         logger.warning(f"Invalid hwnd provided (0). Returning None.")
         return None

    process_id = DWORD(0) # DWORD object to receive the process ID
    thread_id = user32.GetWindowThreadProcessId(hwnd, ctypes.byref(process_id))

    process_id_val = process_id.value # Extract the integer value

    if not thread_id: # GetWindowThreadProcessId returns 0 on failure
        logger.error(f"GetWindowThreadProcessId failed for hwnd={hwnd}. WinError: {ctypes.WinError()}", exc_info=True)
        return None # Failed to get PID

    logger.debug(f"From hwnd={hwnd} got thread_id={thread_id}, process_id={process_id_val}")

    if process_id_val == 0:
         logger.warning(f"Obtained process_id is 0 for hwnd={hwnd}. Returning None.")
         return None

    process_handle = None
    process_name = None

    try:
        # Open the process to query information
        process_handle = kernel32.OpenProcess(PROCESS_ACCESS, False, process_id_val)

        if not process_handle:
            # Check for error code 5 (Access is denied) specifically
            error_code = ctypes.GetLastError()
            if error_code == 5: # ERROR_ACCESS_DENIED
                 logger.warning(f"Access denied when trying to open process (PID: {process_id_val}) for hwnd={hwnd}. May require admin privileges. WinError: {ctypes.WinError()}")
            else:
                 logger.error(f"OpenProcess failed for PID {process_id_val}, hwnd={hwnd}. WinError: {ctypes.WinError()}", exc_info=True)
            return None # Failed to open process

        logger.debug(f"Successfully opened process (PID: {process_id_val}), handle={process_handle}")

        # Get the process executable path
        buffer_size = DWORD(os.pathconf('.', 'PC_PATH_MAX') if hasattr(os, 'pathconf') else 260) # Reasonable buffer size
        filename_buffer = ctypes.create_unicode_buffer(buffer_size.value)

        returned_size = DWORD(0)
        # The path buffer includes the null terminator, so the required size is length + 1.
        # QueryFullProcessImageNameW returns the number of *characters* copied into the buffer, *excluding* the null terminator.
        # So the buffer_size we pass should be the *total size* including space for the null terminator.
        # The returned_size will be the length of the string itself.
        # If the buffer is too small, the function returns 0, GetLastError() returns ERROR_INSUFFICIENT_BUFFER (122),
        # and lpdwSize must be DOUBLED and the call retried.
        # For simplicity here, we use a fixed buffer size and log an error if it's insufficient.
        # A more robust version would handle ERROR_INSUFFICIENT_BUFFER by resizing the buffer.

        success = kernel32.QueryFullProcessImageNameW(
            process_handle,
            0, # dwFlags - 0 for full path
            filename_buffer,
            ctypes.byref(buffer_size) # Pass pointer to buffer size
        )

        if not success:
             error_code = ctypes.GetLastError()
             if error_code == 122: # ERROR_INSUFFICIENT_BUFFER
                 logger.error(f"Buffer too small for process name (PID: {process_id_val}, hwnd={hwnd}). Need to retry with larger buffer. WinError: {ctypes.WinError()}")
             else:
                 logger.error(f"QueryFullProcessImageNameW failed for PID {process_id_val}, hwnd={hwnd}. WinError: {ctypes.WinError()}", exc_info=True)
             return None # Failed to get path

        # success is non-zero on success
        process_path = filename_buffer.value # .value extracts string from buffer

        logger.debug(f"Got process (PID: {process_id_val}) path: {process_path}")

        # Extract the executable file name from the path
        process_name = os.path.basename(process_path)
        logger.info(f"From hwnd={hwnd} (PID: {process_id_val}) got process name: {process_name}")

    except Exception as e:
        logger.error(f"An unexpected error occurred while getting process info for PID {process_id_val}, hwnd={hwnd}: {e}", exc_info=True)
        return None # Ensure None is returned on error
    finally:
        # Always close the process handle if it was successfully opened
        if process_handle and process_handle != ctypes.wintypes.HANDLE(): # Check if handle is valid (not NULL)
            try:
                # Check if handle needs closing (OpenProcess returns a handle that needs closing)
                 if kernel32.CloseHandle(process_handle):
                     logger.debug(f"Closed process handle: {process_handle}")
                 else:
                     # CloseHandle can fail if the handle is invalid, though it shouldn't happen here.
                     logger.warning(f"CloseHandle failed for process handle {process_handle}. WinError: {ctypes.WinError()}")
            except Exception as e:
                 logger.error(f"An error occurred while closing handle {process_handle}: {e}", exc_info=True)

    logger.debug(f"get_process_name_from_hwnd (ctypes) finished, returning: {process_name}")
    return process_name

# Example Usage (for testing this module standalone)
if __name__ == "__main__":
    import sys
    # Adjust sys.path to find logging_config
    project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
    if project_root not in sys.path:
        sys.path.append(project_root)

    from src.utils.logging_config import configure_logging
    # Configure logging for detailed output
    configure_logging(logging.DEBUG)
    logger = logging.getLogger(__name__) # Re-get logger after full configuration

    logger.info("Standalone ctypes process_info test started.")

    # Try to get the current foreground window's process name
    current_foreground_hwnd = user32.GetForegroundWindow() # GetForegroundWindow can also be called via ctypes
    if current_foreground_hwnd:
        print(f"Current foreground window handle (ctypes): {current_foreground_hwnd}")
        process_exe_name = get_process_name_from_hwnd(current_foreground_hwnd)
        if process_exe_name:
            print(f"Current foreground application process name (ctypes): {process_exe_name}")
        else:
            print("Could not get current foreground application process name.")
    else:
        print("Could not get current foreground window handle.")

    # Try getting process name for an invalid handle
    print("\n--- Trying invalid handles (ctypes) ---")
    get_process_name_from_hwnd(0)
    get_process_name_from_hwnd(99999999) # An unlikely handle value

    logger.info("Standalone ctypes process_info test finished.")