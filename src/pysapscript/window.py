from time import sleep

import win32com.client

from pysapscript.types_ import exceptions
from pysapscript.types_.types import NavigateAction
from pysapscript.shell_table import ShellTable


class Window:
    def __init__(
        self,
        connection: int,
        connection_handle: win32com.client.CDispatch,
        session: int,
        session_handle: win32com.client.CDispatch,
    ) -> None:
        self.connection = connection
        self.connection_handle = connection_handle
        self.session = session
        self.session_handle = session_handle

    def __repr__(self) -> str:
        return f"Window(connection={self.connection}, session={self.session})"

    def __str__(self) -> str:
        return f"Window(connection={self.connection}, session={self.session})"

    def __eq__(self, other: object) -> bool:
        if isinstance(other, Window):
            return self.connection == other.connection and self.session == other.session

        return False

    def __hash__(self) -> hash:
        return hash(f"{self.connection_handle}{self.session_handle}")

    def maximize(self) -> None:
        """
        Maximizes this sap window
        """
        self.session_handle.findById("wnd[0]").maximize()

    def restore(self) -> None:
        """
        Restores sap window to its default size, resp. before maximization
        """
        self.session_handle.findById("wnd[0]").restore()

    def close_window(self) -> None:
        """
        Closes this sap window
        """
        self.session_handle.findById("wnd[0]").close()

    def navigate(self, action: NavigateAction) -> None:
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

    def start_transaction(self, transaction: str) -> None:
        """
        Starts transaction

        Args:
            transaction (str): transaction name
        """
        self.write("wnd[0]/tbar[0]/okcd", transaction)
        self.navigate(NavigateAction.enter)

    def press(self, element: str) -> None:
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
            raise exceptions.ActionException(f"Error clicking element {element}: {ex}")

    def select(self, element: str) -> None:
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

    def is_selected(self, element: str) -> bool:
        """
        Gets status of select element

        Args:
            element (str): element select

        Returns:
            bool: selected state

        Raises:
            ActionException: error selecting element

        Example:
            ```
            main_window.is_selected("wnd[2]/tbar[0]/field[1]")
            ```
        """
        try:
            return self.session_handle.findById(element).selected

        except Exception as ex:
            raise exceptions.ActionException(f"Error getting status of element {element}: {ex}")

    def set_checkbox(self, element: str, selected: bool) -> None:
        """
        Selects checkbox element

        Args:
            element (str): checkbox element
            selected (bool): selected state - True for checked, False for unchecked

        Raises:
            ActionException: error checking checkbox element

        Example:
            ```
            main_window.set_checkbox("wnd[0]/usr/chkPA_CHCK", True)
            ```
        """
        try:
            self.session_handle.findById(element).selected = selected

        except Exception as ex:
            raise exceptions.ActionException(f"Error clicking element {element}: {ex}")

    def write(self, element: str, text: str) -> None:
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
            raise exceptions.ActionException(
                f"Error writing to element {element}: {ex}"
            )

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

    def visualize(self, element: str, seconds: int = 1) -> None:
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
            raise exceptions.ActionException(
                f"Error visualizing element {element}: {e}"
            )

    def exists(self, element: str) -> bool:
        """
        checks if element exists by trying to access it

        Args:
            element (str): element to check

        Returns:
            bool: True if element exists, False otherwise
        """

        try:
            self.session_handle.findById(element)
            return True

        except Exception:
            return False

    def send_v_key(
        self,
        element: str = "wnd[0]",
        *,
        focus_element: str | None = None,
        value: int = 0,
    ) -> None:
        """
        Sends VKey to the window, this works for certaion fields
        If more elements are present, optional focus_element can be used to focus on one of them.
        Example use is a pick button, that opens POP-UP window that is otherwise not visible as a separate element.

        Args:
            element (str): element to draw around, default: wnd[0]
            focus_element (str | None): optional, element to focus on, default: None
            value (int): number of VKey to be sent, default: 0

        Raises:
            ActionException: Error focusing or send VKey to element

        Example:
            ```
            window.send_v_key(focus_element="wnd[0]/usr/ctxtCITYC-LOW", value=4)
            ```
        """
        try:
            if focus_element is not None:
                self.session_handle.findById(focus_element).SetFocus()

            self.session_handle.findById(element).sendVKey(value)

        except Exception as e:
            raise exceptions.ActionException(
                f"Error visualizing element {element}: {e}"
            )

    def read_html_viewer(self, element: str) -> str:
        """
        Read the HTML content of the specified HTMLViewer element.

        Parameters:
            element (str): The identifier of the element to read.

        Returns:
            str: The inner HTML content of the specified element.

        Raises:
            ActionException: If an error occurs while reading the element.

        Example:
            ```
            html_content = main_window.read_html_viewer("wnd[0]/usr/cntlGRID1/shellcont[0]/shell")
            ```
        """
        try:
            return self.session_handle.findById(
                element
            ).BrowserHandle.Document.documentElement.innerHTML

        except Exception as e:
            raise exceptions.ActionException(f"Error reading element {element}: {e}")

    def read_shell_table(self, element: str, load_table: bool = True) -> ShellTable:
        """
        Read the table of the specified ShellTable element.
        Args:
            element (str): The identifier of the element to read.
            load_table (bool): Whether to load the table data. Default True

        Returns:
            ShellTable: The ShellTable object with the table data and methods to manage it.

        Example:
            ```
            table = main_window.read_shell_table("wnd[0]/usr/cntlGRID1/shellcont[0]/shell")
            rows = table.rows
            for row in table:
                print(row["COL1"])
            table.to_pandas()
            ```
        """
        return ShellTable(self.session_handle, element, load_table)
