import logging
import os
import threading
import time
import queue
import sys
import ctypes
import ctypes.wintypes
import win32con # Still useful for constants like EVENT_SYSTEM_FOREGROUND, WM_QUIT, etc.

# Import our ctypes helper function to get process name from hwnd
# Ensure the version of process_info.py using ctypes is in src/infrastructure/windows
try:
    from .process_info import get_process_name_from_hwnd
except ImportError as e:
    print(f"Error importing process_info.py: {e}")
    print("Please ensure src/infrastructure/windows/process_info.py exists and uses ctypes.")
    sys.exit(1)

# Ensure logging is configured (ideally done in main.py, but added check here for safety/standalone testing)
logger = logging.getLogger(__name__)
if not logger.handlers:
    # Fallback configuration if not already configured (should not happen if main.py calls logging_config)
    try:
        # Adjust sys.path for standalone testing context if necessary
        project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        if project_root not in sys.path:
             sys.path.append(project_root)

        from src.utils.logging_config import configure_logging
        # Use DEBUG level for detailed event listener logs when testing standalone
        configure_logging(logging.DEBUG)
        logger = logging.getLogger(__name__) # Re-get logger after configuration
        logger.debug("Fallback logging configured in event_listener.py")
    except (ImportError, Exception) as e:
        logging.basicConfig(level=logging.DEBUG, format='[%(asctime)s] [%(levelname)s] [%(name)s.%(funcName)s] - %(message)s')
        logger = logging.getLogger(__name__)
        logger.warning(f"Failed to import/configure custom logging_config. Using basicConfig. Error: {e}")

# --- Define ctypes functions and types needed from user32.dll ---

user32 = ctypes.windll.user32

# Define necessary Windows data types for ctypes
DWORD = ctypes.wintypes.DWORD
HWND = ctypes.wintypes.HWND
HANDLE = ctypes.wintypes.HANDLE
UINT = ctypes.wintypes.UINT
WPARAM = ctypes.wintypes.WPARAM
LPARAM = ctypes.wintypes.LPARAM
LRESULT = ctypes.c_long # Often used as return type for window procedures

# Define the signature for the WinEventProc callback function
# WINEVENTPROC signature:
# void CALLBACK WinEventProc(HWINEVENTHOOK hWinEventHook, DWORD event, HWND hwnd,
#                             LONG idObject, LONG idChild, DWORD dwEventThread,
#                             DWORD dwmsEventTime);
# ctypes.WINFUNCTYPE defines a function prototype for functions exported from DLLs using
# the __stdcall calling convention, which WinEventProc uses.
WinEventProcType = ctypes.WINFUNCTYPE(
    None,         # Return type (void)
    HANDLE,       # hWinEventHook
    DWORD,        # event
    HWND,         # hwnd
    ctypes.c_long,  # idObject
    ctypes.c_long,  # idChild
    DWORD,        # dwEventThread
    DWORD         # dwmsEventTime
)

# Define function signatures for ctypes
# HWINEVENTHOOK SetWinEventHook(DWORD eventMin, DWORD eventMax, ... WinEventProc lpfnWinEventProc, ...)
# SetWinEventHook returns a HWINEVENTHOOK (HANDLE) or 0 on failure.
user32.SetWinEventHook.argtypes = [
    DWORD,      # eventMin
    DWORD,      # eventMax
    HANDLE,     # hmodWinEventProc (handle to the DLL - 0 for WINEVENT_OUTOFCONTEXT)
    WinEventProcType, # lpfnWinEventProc (the callback function pointer)
    DWORD,      # idProcess (0 for all processes)
    DWORD,      # idThread (0 for all threads)
    DWORD       # dwFlags
]
user32.SetWinEventHook.restype = HANDLE # Returns HWINEVENTHOOK (HANDLE)

# BOOL UnhookWinEvent(HWINEVENTHOOK hWinEventHook);
user32.UnhookWinEvent.argtypes = [HANDLE]
user32.UnhookWinEvent.restype = ctypes.wintypes.BOOL # True on success, False on failure

# BOOL GetMessageW(LPMSG lpMsg, HWND hWnd, UINT wMsgFilterMin, UINT wMsgFilterMax);
# For PumpMessages, we process all messages (hWnd=0, filterMin=0, filterMax=0)
MSG = ctypes.wintypes.MSG # Define the MSG structure
LPMSG = ctypes.POINTER(MSG) # Pointer to MSG
user32.GetMessageW.argtypes = [LPMSG, HWND, UINT, UINT]
user32.GetMessageW.restype = ctypes.wintypes.BOOL # Returns non-zero on success, 0 on WM_QUIT, -1 on error

# LRESULT DispatchMessageW(const MSG *lpMsg);
user32.DispatchMessageW.argtypes = [LPMSG]
user32.DispatchMessageW.restype = LRESULT # Return value varies

# BOOL PostThreadMessageW(DWORD idThread, UINT Msg, WPARAM wParam, LPARAM lParam);
user32.PostThreadMessageW.argtypes = [DWORD, UINT, WPARAM, LPARAM]
user32.PostThreadMessageW.restype = ctypes.wintypes.BOOL

# --- Global state for callback communication ---
# When using ctypes for WinAPI callbacks, especially with WINEVENT_OUTOFCONTEXT,
# the callback function (SetWinEventHook's 4th arg) must be a plain C-compatible function
# pointer. This makes it difficult to directly access instance variables (like self._processing_queue).
# A common pattern is to use global/module-level state, managed carefully by the controlling class.
_processing_queue = None       # Global reference to the queue to pass data to main logic
_hook_handle = None            # Global handle to the event hook (needed for Unhook)
_listener_thread_id = None     # Global ID of the thread running the message loop (needed for PostThreadMessage)
_win_event_proc_ref = None     # Global reference to the ctypes callback function pointer,
                               # Important to keep the object alive so it's not GC'd
                               # while the hook is active.

# --- The ctypes WinEventHook callback function ---
# This function MUST match the WinEventProcType signature.
# It is called by the Windows system on the thread that called SetWinEventHook and PumpMessages
@WinEventProcType
def _global_win_event_callback(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
    """
    Global callback function for Windows events, called by OS.
    Matches WINEVENTPROC signature.
    We are primarily interested in EVENT_SYSTEM_FOREGROUND.
    """
    # logger.debug(f"Event received: event={event}, hwnd={hwnd}") # Too verbose for typical use unless debugging callback itself

    if event == win32con.EVENT_SYSTEM_FOREGROUND:
        # This event signifies that a new window has moved to the foreground (received focus).
        # Note: This callback runs on the EventListener thread, not the main application thread.
        # It's crucial this function is fast and does not block.
        logger.debug(f"Foreground window change event received (hwnd={hwnd})")

        # Get the process name using the helper function.
        # This function uses ctypes internally and should be relatively fast.
        process_name = get_process_name_from_hwnd(hwnd)

        if process_name:
            logger.info(f"Foreground activity detected: {process_name}. Submitting to queue.")
            # Pass the process name to the main processing logic via the queue.
            if _processing_queue is not None:
                try:
                    # The `put_nowait` is used to avoid blocking the callback thread
                    # if the queue is full (which shouldn't happen with a large enough queue
                    # or prompt processing in the main thread).
                    _processing_queue.put_nowait(process_name)
                    # logger.debug(f"Submitted '{process_name}' to processing queue.") # Only log if needed for debug
                except queue.Full:
                     logger.warning(f"Processing queue is full, dropped process name: {process_name}")
                except Exception as e:
                    logger.error(f"Failed to put process name into queue: {e}", exc_info=True)
            else:
                 # This indicates a severe setup error or the listener is being called after stop.
                 logger.error("Processing queue is not set in the global callback. Cannot submit process name!", stack_info=True)
        # else: get_process_name_from_hwnd logs warnings/errors if it fails.

# --- Event Listener Class ---
# Manages the listening thread and the Windows Hook using ctypes.
class EventListener:
    """
    Manages the Windows event listening thread for foreground window changes using ctypes.
    """
    def __init__(self, processing_queue: queue.Queue):
        """
        Initializes the EventListener.

        Args:
            processing_queue: A queue.Queue object to send foreground process names to.
        """
        logger.debug("Initializing EventListener instance (ctypes).")
        if not isinstance(processing_queue, queue.Queue):
             logger.error("EventListener requires a queue.Queue instance.", exc_info=True)
             raise TypeError("EventListener requires a queue.Queue instance.")

        self._processing_queue = processing_queue
        self._listener_thread = None
        self._is_running = False
        # We need to store a persistent reference to the ctypes callback function
        # pointer to prevent Python's garbage collector from deleting it
        # while Windows still holds a pointer to it.
        # We'll assign it to a global variable when starting the thread.

    def start(self):
        """
        Starts the event listener thread.
        """
        if self._is_running:
            logger.warning("EventListener is already running.")
            return

        logger.info("Starting EventListener thread (ctypes).")
        self._is_running = True
        # Create the thread, targeting _thread_entry method
        self._listener_thread = threading.Thread(target=self._thread_entry, name="EventListenerThread")
        # Make it non-daemon so main thread can wait for it on stop
        self._listener_thread.daemon = False
        self._listener_thread.start()
        logger.info("EventListener thread started.")

    def stop(self):
        """
        Stops the event listener thread and unhooks the Windows event hook.
        This must be called for clean shutdown.
        """
        if not self._is_running:
            logger.warning("EventListener is not running.")
            # Clear global state even if not running, just in case it was partially started/stuck
            self._clear_global_state()
            return

        logger.info("Stopping EventListener (ctypes).")
        self._is_running = False

        # To stop the message loop (PumpMessages equivalent), we need to post a WM_QUIT message
        # to the message queue of the thread that is running the loop.
        # We stored the thread ID in the global _listener_thread_id.
        global _listener_thread_id
        if _listener_thread_id is not None:
            try:
                 # Use the ctypes PostThreadMessageW function
                success = user32.PostThreadMessageW(_listener_thread_id, win32con.WM_QUIT, 0, 0)
                if success:
                    logger.debug(f"Posted WM_QUIT message to listener thread (ID: {_listener_thread_id}).")
                else:
                     # PostThreadMessage returns 0 if the thread ID is invalid or the queue is full.
                     last_error = ctypes.GetLastError()
                     logger.error(f"PostThreadMessageW failed for thread ID {_listener_thread_id}. Return value: {success}, Last Error: {last_error} ({ctypes.WinError(last_error).strerror})", exc_info=True)

            except Exception as e:
                 logger.error(f"An unexpected error occurred calling PostThreadMessageW: {e}", exc_info=True)
        else:
            logger.warning("Listener thread ID is not available. Cannot post WM_QUIT message.")

        # Wait for the listener thread to finish
        if self._listener_thread and self._listener_thread.is_alive():
            logger.debug("Waiting for EventListener thread (ctypes) to join...")
            self._listener_thread.join(timeout=5.0) # Wait for up to 5 seconds

            if self._listener_thread.is_alive():
                 logger.warning("EventListener thread (ctypes) did not stop within timeout.")
            else:
                 logger.info("EventListener thread (ctypes) joined successfully.")

        # Explicitly clear the global state controlled by this instance *after* joining
        self._clear_global_state()

        logger.info("EventListener stopped (ctypes).")

    def _thread_entry(self):
        """
        The main function executed by the ctypes listener thread.
        Sets up the hook and runs the Windows message loop via ctypes.
        """
        logger.info("EventListener thread entry point (ctypes).")

        # Set the global state BEFORE setting the hook and running the message loop.
        global _processing_queue, _hook_handle, _listener_thread_id, _win_event_proc_ref
        _processing_queue = self._processing_queue
        # Store this thread's ID for posting quit messages
        try:
            # ### 修复 NameError 问题 ###
            # Access GetCurrentThreadId directly via ctypes.windll.kernel32 within the thread
            _listener_thread_id = ctypes.windll.kernel32.GetCurrentThreadId() # <-- 修改为通过 ctypes.windll 访问
            logger.debug(f"Successfully retrieved listener thread ID: {_listener_thread_id}")
        except Exception as e:
            # If even this fails, something is severely wrong
            logger.error(f"FATAL ERROR: Could not get listener thread ID using ctypes.windll.kernel32.GetCurrentThreadId: {e}", exc_info=True)
            # Cannot proceed without a valid thread ID for stopping the message loop later.
            # Indicate that the listener is not running successfully.
            self._is_running = False # Set instance flag to False
            # Note: Clearing global state must now be done carefully or relied on finally block cleanup.
            # Let's attempt immediate cleanup and exit the thread.
            self._cleanup_thread_entry() # Call cleanup logic
            return # Exit the thread function immediately

        # Keep a reference to the ctypes callback function object...
        _win_event_proc_ref = _global_win_event_callback

        try:
            # Set the Windows Event Hook using ctypes
            # ... (user32.SetWinEventHook 调用代码保持不变) ...
            _hook_handle = user32.SetWinEventHook(
                win32con.EVENT_SYSTEM_FOREGROUND,
                win32con.EVENT_SYSTEM_FOREGROUND,
                0,
                _win_event_proc_ref, # 使用全局引用
                0,
                0,
                win32con.WINEVENT_OUTOFCONTEXT | win32con.WINEVENT_SKIPOWNPROCESS
            )
            logger.info(f"Windows Event Hook set successfully (ctypes). Hook handle: {_hook_handle}")

            # Check if the hook was set successfully
            if not _hook_handle: # SetWinEventHook returns NULL handle (0) on failure
                 last_error = ctypes.GetLastError()
                 logger.error(f"Failed to set Windows Event Hook (ctypes)! SetWinEventHook returned 0. Last Error: {last_error} ({ctypes.WinError(last_error).strerror})")
                 # Clean up and exit the thread since hook failed
                 self._cleanup_thread_entry()
                 # No message loop needed if hook failed
                 return

            # Run the Windows message loop.
            # ... (user32.GetMessageW 循环代码保持不变) ...
            logger.info("Entering Windows message loop (ctypes, GetMessageW).")
            msg = MSG()
            while user32.GetMessageW(ctypes.byref(msg), 0, 0, 0):
                 user32.TranslateMessage(ctypes.byref(msg))
                 user32.DispatchMessageW(ctypes.byref(msg))

            logger.info("Windows message loop exited.")

        except Exception as e:
            # Catch any other exceptions during hook setup or message loop
            logger.error(f"An unexpected error occurred in the listener thread (ctypes hook/message loop): {e}", exc_info=True)
        finally:
            # Ensure hook is unhooked and state is cleaned up when the thread exits
            self._cleanup_thread_entry()
            logger.info("EventListener thread finished (ctypes).")

    def _cleanup_thread_entry(self):
        """
        Clean up resources before the listener thread exits.
        Unhooks the Windows event hook and clears global state related to this thread.
        This is called from within the listener thread itself.
        """
        global _hook_handle, _processing_queue, _listener_thread_id, _win_event_proc_ref

        # First, unhook the event if the hook handle is valid.
        if _hook_handle is not None and _hook_handle != 0:
             logger.info(f"Unhooking Windows Event Hook (ctypes): {_hook_handle}")
             try:
                 # Use the ctypes UnhookWinEvent function
                 success = user32.UnhookWinEvent(_hook_handle)
                 if success:
                     logger.info("Windows Event Hook (ctypes) unhooked successfully.")
                 else:
                     # UnhookWinEvent returns FALSE on failure
                      last_error = ctypes.GetLastError()
                      logger.error(f"Failed to unhook Windows Event Hook (ctypes) {_hook_handle}. Return value: {success}, Last Error: {last_error} ({ctypes.WinError(last_error).strerror})", exc_info=True)
             except Exception as e:
                  logger.error(f"An unexpected error occurred calling UnhookWinEvent: {e}", exc_info=True)
             finally:
                 _hook_handle = None # Clear the handle reference

        # Clear the global reference to the ctypes callback function
        # This allows the object to be garbage collected once the hook is removed
        _win_event_proc_ref = None
        logger.debug("Cleared ctypes callback function reference.")

        # Clear the global queue and thread ID references if they match this instance/thread.
        # This helps prevent state from a previous run interfering if cleanup wasn't perfect.
        if _processing_queue is not None and _processing_queue is self._processing_queue:
             _processing_queue = None
             logger.debug("Cleared global processing queue reference.")

        if _listener_thread_id == ctypes.windll.kernel32.GetCurrentThreadId():
             _listener_thread_id = None
             logger.debug("Cleared global listener thread ID reference.")

        logger.debug("EventListener thread cleanup complete (ctypes).")

    def _clear_global_state(self):
        """
        Clears the global state variables associated with this listener instance.
        Called from the main thread's stop() method.
        Note: The actual Unhook and WM_QUIT should be handled *before* calling this,
        ideally by the listener thread itself reacting to WM_QUIT.
        This is a safeguard to release references in the main thread's view.
        """
        global _processing_queue, _hook_handle, _listener_thread_id, _win_event_proc_ref
        # This method's main purpose is to clear references held by the main thread's view.
        # Actual cleanup (Unhook, PumpMessages exit) happens in the listener thread.
        # We should probably verify the listener thread IS NOT _listener_thread_id when calling this for real,
        # to ensure the thread has fully exited. But for simplicity, we clear references here.

        if _processing_queue is self._processing_queue:
             self._processing_queue = None
             _processing_queue = None
             logger.debug("_clear_global_state: Cleared instance and global queue references.")

        # Note: We don't force unhook here. Unhooking should happen in _cleanup_thread_entry
        # after the thread receives WM_QUIT.
        if _hook_handle is not None: # Just log if hook handle is still set unexpectedly
             logger.warning(f"_clear_global_state: Global hook handle {_hook_handle} was still set.")
             _hook_handle = None # Force clear global reference

        if _listener_thread_id is not None:
            logger.warning(f"_clear_global_state: Global listener thread ID {_listener_thread_id} was still set.")
            _listener_thread_id = None # Force clear global reference

        if _win_event_proc_ref is not None:
             logger.warning("_clear_global_state: Global ctypes callback reference was still set.")
             _win_event_proc_ref = None # Force clear global reference

# --- Example Usage (for testing this module standalone) ---
# In a real application, this would be managed by main.py and PowerSwitcherApp.
if __name__ == "__main__":
    # Log configuration is handled by the fallback at the top or main program.
    logger.info("Standalone EventListener (ctypes) test started.")
    logger.info("Try switching foreground windows. Press Ctrl+C to stop.")

    # Create a queue to receive process names from the callback
    event_queue = queue.Queue()

    # Create and start the listener
    listener = EventListener(event_queue)
    listener.start()

    # --- Main thread logic (simplified for this example) ---
    # In the real app, this would be the main loop (like the GUI loop) or a processing loop.
    # Here, we'll just periodically check the queue for incoming process names.
    try:
        # Keep the main thread alive and responsive to queue and KeyboardInterrupt
        while True:
            try:
                # Get a process name from the queue with a short timeout.
                # The timeout allows the main thread to periodically check for KeyboardInterrupt.
                process_name = event_queue.get(timeout=0.1)
                logger.info(f"Main thread received PROCESS from queue: {process_name}")
                # === PLACEHOLDER for Power Switching Logic ===
                # In the real application, you would call your power switching
                # module here based on the received process_name.
                # For this test, we just log it.
                # =============================================

            except queue.Empty:
                # No items in the queue, just continue looping and checking for interrupt
                pass
            except KeyboardInterrupt:
                logger.info("Keyboard interrupt detected in main thread. Stopping listener.")
                break # Exit the loop on Ctrl+C
            except Exception as e:
                logger.error(f"An unexpected error occurred in main thread queue handling loop: {e}", exc_info=True)
                break # Exit on other errors
            # Small sleep to yield CPU, though the queue.get(timeout) already provides some yielding.
            # time.sleep(0.01)

    finally:
        # Cleanly stop the listener thread and unhook before the main program exits
        logger.info("Main thread cleanup: stopping listener.")
        listener.stop()
        # Ensure the listener thread has finished before the main thread exits
        if listener._listener_thread and listener._listener_thread.is_alive():
             logger.info("Waiting for listener thread to finish after stop signal.")
             listener._listener_thread.join() # Wait without timeout
        logger.info("Standalone EventListener (ctypes) test finished.")