import os
import time
import atexit
from pathlib import Path
from subprocess import Popen
from typing import NamedTuple

import pandas
import win32com.client
from win32gui import FindWindow, GetWindowText



class WindowDidNotAppearException(Exception):
    """Main windows didn't show up - possible pop-up window"""


class AttachException(Exception):
    """Error with attaching - connection or session"""


class ActionException(Exception):
    """Error performing action - click, select ..."""


class Obj(NamedTuple):
    connection: int
    session: int


class Sap:
    def __init__(self):
        self.sap_gui_auto = None
        self.application = None
        self.connections = {}
        self.sessions = {}

        atexit.register(self._quit)

    def start_sap(self, sap_dir: Path, sid: str, client: str, user: str, password: str) -> Obj:
        self._launch(
            sap_dir, sid, client, user, password
        )
        obj: Obj = self.attach(0, 0)
        self.maximize(obj)
        time.sleep(5)

        return obj

    def _quit(self):
        """tries to close the sap normal way (from main wnd) and then kills the process"""

        try:
            main = self.attach(0, 0)
            self.element_select(main, "wnd[0]/mbar/menu[4]/menu[12]")
            self.element_press(main, "wnd[1]/usr/btnSPOP-OPTION1")

        finally:
            self.kill_process("saplogon.exe")

        self.sap_gui_auto = None
        self.application = None
        self.connections = {}
        self.sessions = {}


    def attach(self, connection: int, session: int) -> Obj:
        """connects to sap gui by number, sets application,
        connection and session, then returns int list of conn and sess"""

        if not str(connection).isnumeric():
            raise Exception("Wrong connection argument!")

        if not str(session).isnumeric():
            raise Exception("Wrong session argument!")

        if not isinstance(self.sap_gui_auto, win32com.client.CDispatch):
            self.sap_gui_auto = win32com.client.GetObject("SAPGUI")

        if not isinstance(self.application, win32com.client.CDispatch):
            self.application = self.sap_gui_auto.GetScriptingEngine

        if connection not in self.connections:
            try:
                self.connections[connection] = self.application.Children(connection)
                self.sessions[connection] = {}

            except Exception:
                raise AttachException("Could not attach connection %s!" % connection)

        if session not in self.sessions[connection]:
            try:
                self.sessions[connection][session] = self.connections[
                    connection
                ].Children(session)

            except Exception:
                raise AttachException("Could not attach session %s!" % session)

        return Obj(connection, session)

    def detach(self, obj: Obj):
        """sets connection or/and session to none"""

        if str(type(obj)) != "<class 'list'>":
            raise Exception("Wrong object argument!")

        self.sessions[obj.connection].pop(obj.session, None)

        if len(self.sessions[obj.connection]) == 0:
            self.sessions.pop(obj.connection, None)
            self.connections.pop(obj.connection, None)

        return

    @staticmethod
    def wait_for_window_title(title: str, timeout: int = 10):
        """loops until title of window appears"""

        c = 0
        while True:
            if c > timeout:
                raise WindowDidNotAppearException(
                    "Window title %s didn't appear within time window!" % title
                )

            window_pid = FindWindow("SAP_FRONTEND_SESSION", None)
            window_text = GetWindowText(window_pid)
            if window_text.startswith(title):
                break

            time.sleep(1)
            c += 1

    @staticmethod
    def kill_process(process: str):
        os.system("taskkill /f /im %s" % process)

    def _launch(self, working_dir: Path, sid: str, client: str, user: str, password: str):
        """launches sap from sapshcut.exe"""

        working_dir = working_dir.resolve()
        sap_executable = working_dir.joinpath("sapshcut.exe")
        command = f"-system={sid} -client={client} -user={user} -pw={password}"
        tryouts = 2

        while True:
            try:
                Popen([sap_executable, *command.split(" ")])

                self.wait_for_window_title("SAP Easy Access")
                break

            except WindowDidNotAppearException:
                tryouts = tryouts - 1
                self.kill_process("saplogon.exe")

                if tryouts == 0:
                    raise WindowDidNotAppearException(
                        "Failed to launch SAP and Mindow did not appear."
                    )
        return

    def maximize(self, obj: Obj):
        """Maximizes sap window"""

        self.sessions[obj.connection][obj.session].findById("wnd[0]").maximize()

    def open_new_window(self):
        """Opens new sap window, first session must be already set up"""

        if 0 in self.sessions and 0 in self.sessions[0]:
            self.sessions[0][0].createSession()

        # this window title wait method will only work if main sap window doesn't have same title
        self.wait_for_window_title("SAP Easy Access")

    def navigate(self, obj: Obj, action: str):
        """Navigates SAP: enter, back, end, cancel, save"""

        if action == "enter":
            el = "wnd[0]/tbar[0]/btn[0]"
        elif action == "back":
            el = "wnd[0]/tbar[0]/btn[3]"
        elif action == "end":
            el = "wnd[0]/tbar[0]/btn[15]"
        elif action == "cancel":
            el = "wnd[0]/tbar[0]/btn[12]"
        elif action == "save":
            el = "wnd[0]/tbar[0]/btn[13]"
        else:
            raise ActionException("Wrong navigation action!")

        self.sessions[obj.connection][obj.session].findById(el).press()

    def start_transaction(self, obj: Obj, transaction: str):
        self.element_write(obj, "wnd[0]/tbar[0]/okcd", transaction)
        self.navigate(obj, "enter")

    def element_press(self, obj: Obj, element: str):
        """Presses element"""

        try:
            self.sessions[obj.connection][obj.session].findById(element).press()

        except Exception as ex:
            raise ActionException(f"Error clicking element {element}: {ex}")

    def element_select(self, obj: Obj, element: str):
        """Presses element"""

        try:
            self.sessions[obj.connection][obj.session].findById(element).select()

        except Exception as ex:
            raise ActionException(f"Error clicking element {element}: {ex}")

    def element_write(self, obj: Obj, element: str, text: str):
        """Sets property text to a value"""

        try:
            self.sessions[obj.connection][obj.session].findById(element).text = text

        except Exception as ex:
            raise ActionException(f"Error writing to element {element}: {ex}")

    def element_read(self, obj: Obj, element: str) -> str:
        """Reads property text"""

        try:
            return self.sessions[obj.connection][obj.session].findById(element).text

        except Exception as e:
            raise ActionException(f"Error reading element {element}: {e}")

    def shell_table_read(self, obj: Obj, element: str) -> pandas.DataFrame:
        """Reads table of shell table and returns pandas DataFrame"""

        try:
            shell = self.sessions[obj.connection][obj.session].findById(element)

            columns = shell.ColumnOrder
            c_rows = shell.RowCount

            df = pandas.DataFrame()

            for i in range(0, c_rows):
                df_to_append = {}
                for col in columns:
                    df_to_append[col] = shell.GetCellValue(i, col)

                df = pandas.concat(
                    [df, pandas.DataFrame(df_to_append, index=[0])], ignore_index=True
                )

            return df

        except Exception as ex:
            raise ActionException(f"Error reading element {element}: {ex}")

    def shell_button_press(self, obj: Obj, element: str, button: str):
        self.sessions[obj.connection][obj.session].findById(element).pressButton(button)

    def shell_checkbox_change(self, obj: Obj, element: str, checkbox: str, flag: bool):
        self.sessions[obj.connection][obj.session].findById(element).changeCheckbox(checkbox, "1", flag)
