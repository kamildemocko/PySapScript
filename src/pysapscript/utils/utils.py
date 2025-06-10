import os
import time

from win32gui import FindWindow, GetWindowText

from pysapscript.types_.exceptions import WindowDidNotAppearException


def kill_process(process: str):
    """
    Kills process by process name

    Args:
        process (str): process name
    """
    os.system("taskkill /f /im %s" % process)


def wait_for_window_title(title: str, timeout_loops: int = 30):
    """
    loops until title of expected window appears,
    waits for 1 second between each check

    Args:
        title (str): expected window title
        timeout_loops (int): number of loops

    Raises:
        WindowDidNotAppearException: Expected window did not appear
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
