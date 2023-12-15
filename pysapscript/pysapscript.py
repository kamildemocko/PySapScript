import time
import atexit
from pathlib import Path
from subprocess import Popen

import win32com.client

from pysapscript import window
from pysapscript.utils import utils
from pysapscript.types_ import exceptions


class Sapscript:
    def __init__(self, default_window_title: str = "SAP Easy Access"):
        self.sap_gui_auto = None
        self.application = None
        self.default_window_title = default_window_title

    def launch_sap(self, *, root_sap_dir: Path = Path(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui"), 
                   sid: str, client: str, 
                   user: str, password: str, 
                   maximise: bool = True, quit_auto: bool = True):
        """
        Launches SAP and waits for it to load, 
        quit_auto: quits automatically on exit if set to True
        """

        self._launch(
            root_sap_dir, 
            sid, 
            client, 
            user, 
            password, 
            maximise
        )

        time.sleep(5)
        if quit_auto:
            atexit.register(self.quit)

    def quit(self):
        """
        Tries to close the sap normal way (from main wnd) 
        Then kills the process
        """

        try:
            main_window = self.attach_window(0, 0)
            main_window.select("wnd[0]/mbar/menu[4]/menu[12]")
            main_window.press("wnd[1]/usr/btnSPOP-OPTION1")

        finally:
            utils.kill_process("saplogon.exe")

    def attach_window(self, connection: int, session: int) -> window.Window:
        """
        Attaches window by connection and session number
        Connection and session start with 0
        """

        if not isinstance(connection, int):
            raise Exception("Wrong connection argument!")

        if not isinstance(session, int):
            raise Exception("Wrong session argument!")

        if not isinstance(self.sap_gui_auto, win32com.client.CDispatch):
            self.sap_gui_auto = win32com.client.GetObject("SAPGUI")

        if not isinstance(self.application, win32com.client.CDispatch):
            self.application = self.sap_gui_auto.GetScriptingEngine

        try:
            connection_handle = self.application.Children(connection)

        except Exception:
            raise exceptions.AttachException("Could not attach connection %s!" % connection)

        try:
            session_handle = connection_handle.Children(session)

        except Exception:
            raise exceptions.AttachException("Could not attach session %s!" % session)

        return window.Window(
            connection=connection,
            connection_handle=connection_handle,
            session=session,
            session_handle=session_handle,
        )

    def open_new_window(self, window_to_handle_opening: window.Window):
        """
        Opens new sap window, 
        First session must be already set up
        """

        window_to_handle_opening.session_handle.createSession()

        utils.wait_for_window_title(self.default_window_title)

    def _launch(self, working_dir: Path, sid: str, client: str, 
                user: str, password: str, maximise: bool):
        """launches sap from sapshcut.exe"""

        working_dir = working_dir.resolve()
        sap_executable = working_dir.joinpath("sapshcut.exe")

        maximise_sap = "-max" if maximise else ""
        command = f"-system={sid} -client={client} "\
            f"-user={user} -pw={password} {maximise_sap}"

        tryouts = 2
        while tryouts > 0:
            try:
                Popen([sap_executable, *command.split(" ")])

                utils.wait_for_window_title(self.default_window_title)
                break

            except exceptions.WindowDidNotAppearException:
                tryouts = tryouts - 1
                utils.kill_process("saplogon.exe")

        else:
            raise exceptions.WindowDidNotAppearException(
                "Failed to launch SAP - Mindow did not appear."
            )
