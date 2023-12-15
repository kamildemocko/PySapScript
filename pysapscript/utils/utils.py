import os
import time

from win32gui import FindWindow, GetWindowText

from pysapscript.types_.exceptions import WindowDidNotAppearException

def kill_process(process: str):
    os.system("taskkill /f /im %s" % process)

def wait_for_window_title(title: str, timeout_loops: int = 10):
    """
    loops until title of window appears, 
    waits for 1 second between checks
    """

    for _ in range(0, timeout_loops):

        window_pid = FindWindow("SAP_FRONTEND_SESSION", None)
        window_text = GetWindowText(window_pid)
        if window_text.startswith(title):
            break

        time.sleep(1)

    else:
        raise WindowDidNotAppearException(
            "Window title %s didn't appear within time window!" % title
        )
