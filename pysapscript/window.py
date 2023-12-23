from time import sleep

import win32com.client
import pandas
from win32com.universal import com_error

from pysapscript.types_ import exceptions
from pysapscript.types_.types import NavigateAction


class Window:
    def __init__(self, 
                 connection: int, connection_handle: win32com.client.CDispatch, 
                 session: int, session_handle: win32com.client.CDispatch):
        self.connection = connection
        self.connection_handle = connection_handle
        self.session = session
        self.session_handle = session_handle

    def maximize(self):
        """
        Maximizes this sap window
        """

        self.session_handle.findById("wnd[0]").maximize()

    def restore(self):
        """
        Restores sap window to its default size, resp. before maximization
        """

        self.session_handle.findById("wnd[0]").restore()

    def close_window(self):
        """
        Closes this sap window
        """

        self.session_handle.findById("wnd[0]").close()

    def navigate(self, action: NavigateAction):
        """
        Navigates SAP: enter, back, end, cancel, save

        Args:
            action (NavigateAction): enter, back, end, cancel, save

        Raises:
            ActionException: wrong navigation action

        Example:
            ```
            main_window.navigate(NavigateAction.enter)
            ```
        """

        if action == NavigateAction.enter:
            el = "wnd[0]/tbar[0]/btn[0]"
        elif action == NavigateAction.back:
            el = "wnd[0]/tbar[0]/btn[3]"
        elif action == NavigateAction.end:
            el = "wnd[0]/tbar[0]/btn[15]"
        elif action == NavigateAction.cancel:
            el = "wnd[0]/tbar[0]/btn[12]"
        elif action == NavigateAction.save:
            el = "wnd[0]/tbar[0]/btn[13]"
        else:
            raise exceptions.ActionException("Wrong navigation action!")

        self.session_handle.findById(el).press()

    def start_transaction(self, transaction: str):
        """
        Starts transaction

        Args:
            transaction (str): transaction name
        """

        self.write("wnd[0]/tbar[0]/okcd", transaction)
        self.navigate(NavigateAction.enter)

    def press(self, element: str):
        """
        Presses element

        Args:
            element (str): element to press

        Raises:
            ActionException: error clicking element

        Example:
            ```
            main_window.press("wnd[0]/usr/tabsTABTC/tabxTAB03/subIncl/SAPML03")
            ```
        """

        try:
            self.session_handle.findById(element).press()

        except Exception as ex:
            raise exceptions.ActionException(
                f"Error clicking element {element}: {ex}"
            )

    def select(self, element: str):
        """
        Selects element or menu item

        Args:
            element (str): element to select - tabs, menu items

        Raises:
            ActionException: error selecting element

        Example:
            ```
            main_window.select("wnd[2]/tbar[0]/btn[1]")
            ```
        """

        try:
            self.session_handle.findById(element).select()

        except Exception as ex:
            raise exceptions.ActionException(f"Error clicking element {element}: {ex}")

    def write(self, element: str, text: str):
        """
        Sets text property of an element

        Args:
            element (str): element to accept a value
            text (str): value to set

        Raises:
            ActionException: Error writing to element

        Example:
            ```
            main_window.write("wnd[0]/usr/tabsTABTC/tabxTAB03/subIncl/SAPML03", "VALUE")
            ```
        """

        try:
            self.session_handle.findById(element).text = text

        except Exception as ex:
            raise exceptions.ActionException(f"Error writing to element {element}: {ex}")

    def read(self, element: str) -> str:
        """
        Reads text property

        Args:
            element (str): element to read

        Raises:
            ActionException: Error reading element

        Example:
            ```
            value = main_window.read("wnd[0]/usr/tabsTABTC/tabxTAB03/subIncl/SAPML03")
            ```
        """

        try:
            return self.session_handle.findById(element).text

        except Exception as e:
            raise exceptions.ActionException(f"Error reading element {element}: {e}")

    def visualize(self, element: str, seconds: int = 1):
        """
        draws red frame around the element

        Args:
            element (str): element to draw around
            seconds (int): seconds to wait for

        Raises:
            ActionException: Error visualizing element
        """

        try:
            self.session_handle.findById(element).Visualize(1)
            sleep(seconds)

        except Exception as e:
            raise exceptions.ActionException(f"Error visualizing element {element}: {e}")

    def read_shell_table(self, element: str, load_table: bool = True) -> pandas.DataFrame:
        """
        Reads table of shell table

        If the table is too big, the SAP will not render all the data.
        Default is to load table before reading it

        Args:
            element (str): table element
            load_table (bool): whether to load table before reading

        Returns:
            pandas.DataFrame: table data

        Raises:
            ActionException: Error reading table

        Example:
            ```
            table = main_window.read_shell_table("wnd[0]/usr/shellContent/shell")
            ```
        """

        try:
            shell = self.session_handle.findById(element)

            columns = shell.ColumnOrder
            rows_count = shell.RowCount

            if rows_count == 0:
                return pandas.DataFrame()

            if load_table:
                self.load_shell_table(element)

            data = [{column: shell.GetCellValue(i, column) for column in columns} for i in range(rows_count)]

            return pandas.DataFrame(data)

        except Exception as ex:
            raise exceptions.ActionException(f"Error reading element {element}: {ex}")

    def load_shell_table(self, table_element: str, move_by: int = 20, move_by_table_end: int = 2):
        """
        Skims through the table to load all data, as SAP only loads visible data

        Args:
            table_element (str): table element
            move_by (int): number of rows to move by, default 20
            move_by_table_end (int): number of rows to move by when reaching the end of the table, default 2

        Raises:
            ActionException: error finding table
        """

        row_position = 0

        try:
            shell = self.session_handle.findById(table_element)

        except Exception as e:
            raise exceptions.ActionException(f"Error finding table {table_element}: {e}")

        while True:
            try:
                shell.currentCellRow = row_position
                shell.SelectedRows = row_position

            except com_error:
                """no more rows for this step"""
                break

            row_position += move_by

        row_position -= 20
        while True:
            try:
                shell.currentCellRow = row_position
                shell.SelectedRows = row_position

            except com_error:
                """no more rows for this step"""
                break

            row_position += move_by_table_end

    def press_shell_button(self, element: str, button: str):
        """
        Presses button that is in a shell table

        Args:
            element (str): table element
            button (str): button name

        Raises:
            ActionException: error pressing shell button

        Example:
            ```
            main_window.press_shell_button("wnd[0]/usr/shellContent/shell", "%OPENDWN")
            ```
        """

        try:
            self.session_handle.findById(element).pressButton(button)

        except Exception as e:
            raise exceptions.ActionException(f"Error pressing button {button}: {e}")

    def change_shell_checkbox(self, element: str, checkbox: str, flag: bool):
        """
        Sets checkbox in a shell table

        Args:
            element (str): table element
            checkbox (str): checkbox name
            flag (bool): True for checked, False for unchecked

        Raises:
            ActionException: error setting shell checkbox

        Example:
            ```
            main_window.change_shell_checkbox("wnd[0]/usr/cntlALV_CONT/shellcont/shell/rows[1]", "%CHBX", True)
            ```
        """

        try:
            self.session_handle.findById(element).changeCheckbox(checkbox, "1", flag)

        except Exception as e:
            raise exceptions.ActionException(f"Error setting element {element}: {e}")
