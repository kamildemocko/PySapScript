import win32com.client
import pandas

from pysapscript.src.types import exceptions
from pysapscript.src.types.types import NavigateAction


class Window:
    def __init__(self, 
                 connection: int, connection_handle: win32com.client.CDispatch, 
                 session: int, session_handle: win32com.client.CDispatch):
        self.connection = connection
        self.connection_handle = connection_handle
        self.session = session
        self.session_handle = session_handle

    def maximize(self):
        """Maximizes this sap window"""

        self.session_handle.findById("wnd[0]").maximize()

    def navigate(self, action: NavigateAction):
        """Navigates SAP: enter, back, end, cancel, save"""

        match action:
            case NavigateAction.enter:
                el = "wnd[0]/tbar[0]/btn[0]"
            case NavigateAction.back:
                el = "wnd[0]/tbar[0]/btn[3]"
            case NavigateAction.end:
                el = "wnd[0]/tbar[0]/btn[15]"
            case NavigateAction.cancel:
                el = "wnd[0]/tbar[0]/btn[12]"
            case NavigateAction.save:
                el = "wnd[0]/tbar[0]/btn[13]"
            case _:
                raise exceptions.ActionException("Wrong navigation action!")

        self.session_handle.findById(el).press()

    def start_transaction(self, transaction: str):
        self.write("wnd[0]/tbar[0]/okcd", transaction)
        self.navigate(NavigateAction.enter)

    def press(self, element: str):
        """Presses element"""

        try:
            self.session_handle.findById(element).press()

        except Exception as ex:
            raise exceptions.ActionException(
                f"Error clicking element {element}: {ex}"
            )

    def select(self, element: str):
        """Presses element"""

        try:
            self.session_handle.findById(element).select()

        except Exception as ex:
            raise exceptions.ActionException(f"Error clicking element {element}: {ex}")

    def write(self, element: str, text: str):
        """Sets property text to a value"""

        try:
            self.session_handle.findById(element).text = text

        except Exception as ex:
            raise exceptions.ActionException(f"Error writing to element {element}: {ex}")

    def read(self, element: str) -> str:
        """Reads property text"""

        try:
            return self.session_handle.findById(element).text

        except Exception as e:
            raise exceptions.ActionException(f"Error reading element {element}: {e}")

    def read_shell_table(self, element: str) -> pandas.DataFrame:
        """Reads table of shell table and returns pandas DataFrame"""

        try:
            shell = self.session_handle.findById(element)

            columns = shell.ColumnOrder
            rows_count = shell.RowCount

            data = [{column: shell.GetCellValue(i, column) for column in columns} for i in range(rows_count)]

            return pandas.DataFrame(data)

        except Exception as ex:
            raise exceptions.ActionException(f"Error reading element {element}: {ex}")

    def press_shell_button(self, element: str, button: str):
        self.session_handle.findById(element).pressButton(button)

    def change_shell_checkbox(self, element: str, checkbox: str, flag: bool):
        self.session_handle.findById(element).changeCheckbox(checkbox, "1", flag)

    def close_window(self):
        """Closes this sap window, assuming it's on the first page"""

        self.session_handle.findById("wnd[0]").close()
