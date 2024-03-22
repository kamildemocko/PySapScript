import time
import atexit
from pathlib import Path
from subprocess import Popen

import win32com.client

from pysapscript import window
from pysapscript.utils import utils
from pysapscript.types_ import exceptions


class Sapscript:
    def __init__(self, default_window_title: str = "SAP Easy Access") -> None:
        """
        Args:
            default_window_title (str): default SAP window title

        Example:
            sapscript = Sapscript()
            main_window = sapscript.attach_window(0, 0)
            main_window.write("wnd[0]/tbar[0]/okcd", "ZLOGON")
            main_window.press("wnd[0]/tbar[0]/btn[0]")
        """
        self._sap_gui_auto = None
        self._application = None
        self.default_window_title = default_window_title

    def __repr__(self) -> str:
        return f"Sapscript(default_window_title={self.default_window_title})"

    def __str__(self) -> str:
        return f"Sapscript(default_window_title={self.default_window_title})"

    def launch_sap(self, *, root_sap_dir: Path = Path(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui"), 
                   sid: str, client: str, 
                   user: str, password: str, 
                   maximise: bool = True, quit_auto: bool = True) -> None:
        """
        Launches SAP and waits for it to load

        Args:
            root_sap_dir (pathlib.Path): SAP directory in the system
            sid (str): SAP system ID
            client (str): SAP client
            user (str): SAP user
            password (str): SAP password
            maximise (bool): maximises window after start if True
            quit_auto (bool): quits automatically on SAP exit if True

        Raises:
            WindowDidNotAppearException: No SAP window appeared

        Example:
            ```
            pss.launch_sap(
                sid="SQ4",
                client="012",
                user="robot_t",
                password=os.getenv("secret")
            )
            ```
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

    def quit(self) -> None:
        """
        Tries to close the sap normal way (from main window), then kills the process
        """
        try:
            main_window = self.attach_window(0, 0)
            main_window.select("wnd[0]/mbar/menu[4]/menu[12]")
            main_window.press("wnd[1]/usr/btnSPOP-OPTION1")

        finally:
            utils.kill_process("saplogon.exe")

    def attach_window(self, connection: int, session: int) -> window.Window:
        """
        Attaches window by connection and session number ID

        Connection starts with 0 and is +1 for each client
        Session start with 0 and is +1 for each new window opened

        Args:
            connection (int): connection number
            session (int): session number

        Returns:
            window.Window: window attached

        Raises:
            AttributeError: srong connection or session
            AttachException: could not attach to SAP window

        Example:
            ```
            main_window = pss.attach_window(0, 0)
            ```
        """
        if not isinstance(connection, int):
            raise AttributeError("Wrong connection argument!")

        if not isinstance(session, int):
            raise AttributeError("Wrong session argument!")

        if not isinstance(self._sap_gui_auto, win32com.client.CDispatch):
            self._sap_gui_auto = win32com.client.GetObject("SAPGUI")

        if not isinstance(self._application, win32com.client.CDispatch):
            self._application = self._sap_gui_auto.GetScriptingEngine

        try:
            connection_handle = self._application.Children(connection)

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

    def open_new_window(self, window_to_handle_opening: window.Window) -> None:
        """
        Opens new sap window

        SAP must be already launched and window that is not busy must be available
        Warning, as of now, this method will not wait for window to appear if any
        other window is opened and has the default windows title

        Args:
            window_to_handle_opening: idle SAP window that will be used to open new window

        Raises:
            WindowDidNotAppearException: no SAP window appeared

        Example:
            ```
            main_window = pss.attach_window(0, 0)
            pss.open_new_window(main_window)
            ```
        """
        window_to_handle_opening.session_handle.createSession()

        utils.wait_for_window_title(self.default_window_title)

    def _launch(self, working_dir: Path, sid: str, client: str, 
                user: str, password: str, maximise: bool) -> None:
        """
        launches sap from sapshcut.exe
        """

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
