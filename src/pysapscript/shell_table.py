from typing import Self, Any
from typing import overload

import win32com.client
from win32com.universal import com_error
import polars as pl
import pandas

from pysapscript.types_ import exceptions


class ShellTable:
    """
    A class representing a shell table
    """

    def __init__(self, session_handle: win32com.client.CDispatch, element: str, load_table: bool = True) -> None:
        """
        Usually table contains a table of data, but it can also be a non-data shell table, that holds toolbar

        Args:
            session_handle (win32com.client.CDispatch): SAP session handle
            element (str): SAP table element
            load_table (bool): loads table if True, default True

        Raises:
            ActionException: error reading table data
        """
        self.table_element = element
        self._session_handle = session_handle
        self.data_present = False
        self.data = self._read_shell_table(load_table)
        self.rows = self.data.shape[0]
        self.columns = self.data.shape[1]

    def __repr__(self) -> str:
        return repr(self.data)

    def __str__(self) -> str:
        return str(self.data)

    def __eq__(self, other: object) -> bool:
        if isinstance(other, ShellTable):
            return self.data.equals(other.data)
        else:
            raise NotImplementedError(f"Cannot compare ShellTable with {type(other)}")

    def __hash__(self) -> int:
        return hash(f"{self._session_handle}{self.table_element}{self.data.shape}")

    def __getitem__(self, item: object) -> dict[str, Any] | list[dict[str, Any]]:
        if self.data_present is False:
            raise ValueError("Data was not found in shell table")

        if isinstance(item, int):
            return self.data.row(item, named=True)
        elif isinstance(item, slice):
            if item.step is not None:
                raise NotImplementedError("Step is not supported")

            sl = self.data.slice(item.start, item.stop - item.start)
            return sl.to_dicts()
        else:
            raise ValueError("Incorrect type of index")

    def __iter__(self) -> "ShellTableRowIterator":
        return ShellTableRowIterator(self.data)

    def _read_shell_table(self, load_table: bool = True) -> pl.DataFrame:
        """
        Reads table of shell table

        If the table is too big, the SAP will not render all the data.
        Default is to load table before reading it

        Args:
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
            shell = self._session_handle.findById(self.table_element)

            if hasattr(shell, "ColumnOrder") is False or hasattr(shell, "RowCount") is False:
                return pl.DataFrame()

            columns = shell.ColumnOrder
            rows_count = shell.RowCount

            if rows_count == 0:
                return pl.DataFrame()

            self.data_present = True

            if load_table:
                self.load()

            data = [
                {column: shell.GetCellValue(i, column) for column in columns}
                for i in range(rows_count)
            ]

            return pl.DataFrame(data)

        except Exception as ex:
            raise exceptions.ActionException(f"Error reading element {self.table_element}: {ex}")

    def to_polars_dataframe(self) -> pl.DataFrame:
        """
        Get table data as a polars DataFrame

        Returns:
            polars.DataFrame: table data
        """
        if self.data_present is False:
            raise ValueError("Data was not found in shell table")

        return self.data

    def to_pandas_dataframe(self) -> pandas.DataFrame:
        """
        Get table data as a pandas DataFrame

        Returns:
            pandas.DataFrame: table data
        """
        if self.data_present is False:
            raise ValueError("Data was not found in shell table")

        return self.data.to_pandas()

    def to_dict(self) -> dict[str, Any]:
        """
        Get table data as a dictionary

        Returns:
            dict[str, Any]: table data in named dictionary - column names as keys
        """
        if self.data_present is False:
            raise ValueError("Data was not found in shell table")
        
        return self.data.to_dict(as_series=False)

    def to_dicts(self) -> list[dict[str, Any]]:
        """
        Get table data as a list of dictionaries

        Returns:
            list[dict[str, Any]]: table data in list of named dictionaries - rows
        """
        if self.data_present is False:
            raise ValueError("Data was not found in shell table")

        return self.data.to_dicts()

    def get_column_names(self) -> list[str]:
        """
        Get column names

        Returns:
            list[str]: column names in the table
        """
        if self.data_present is False:
            raise ValueError("Data was not found in shell table")

        return self.data.columns

    @overload
    def cell(self, row: int, column: int) -> Any:
        ...

    @overload
    def cell(self, row: int, column: str) -> Any:
        ...

    def cell(self, row: int, column: str | int) -> Any:
        """
        Get cell value

        Args:
            row (int): row index
            column (str | int): column name or index

        Returns:
            Any: cell value
        """
        if self.data_present is False:
            raise ValueError("Data was not found in shell table")

        return self.data.item(row, column)

    def load(self, move_by: int = 20, move_by_table_end: int = 2) -> None:
        """
        Skims through the table to load all data, as SAP only loads visible data

        Args:
            move_by (int): number of rows to move by, default 20
            move_by_table_end (int): number of rows to move by when reaching the end of the table, default 2

        Raises:
            ActionException: error finding table
        """
        if self.data_present is False:
            raise ValueError("Data was not found in shell table")

        row_position = 0

        try:
            shell = self._session_handle.findById(self.table_element)

        except Exception as e:
            raise exceptions.ActionException(
                f"Error finding table {self.table_element}: {e}"
            )

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

    def press_button(self, button: str) -> None:
        """
        Presses button that is in a shell table

        Args:
            button (str): button name

        Raises:
            ActionException: error pressing shell button

        Example:
            ```
            main_window.press_shell_button("wnd[0]/usr/shellContent/shell", "%OPENDWN")
            ```
        """
        try:
            self._session_handle.findById(self.table_element).pressButton(button)

        except Exception as e:
            raise exceptions.ActionException(f"Error pressing button {button}: {e}")

    def select_rows(self, indexes: list[int]) -> None:
        """
        Selects rows (visual) is in a shell table

        Args:
            indexes (list[int]): indexes of rows to select, starting with 0

        Raises:
            ActionException: error selecting shell rows

        Example:
            ```
            main_window.select_shell_rows("wnd[0]/usr/shellContent/shell", [0, 1, 2])
            ```
        """
        if self.data_present is False:
            raise ValueError("Data was not found in shell table")

        try:
            value = ",".join([str(n) for n in indexes])
            self._session_handle.findById(self.table_element).selectedRows = value

        except Exception as e:
            raise exceptions.ActionException(
                f"Error selecting rows with indexes {indexes}: {e}"
            )

    def select_row(self, index: int) -> None:
        """
        Selects row and set it as active in a shell table

        Args:
            indexes (int): indexe of row to select, starting with 0

        Raises:
            ActionException: error selecting shell row

        Example:
            ```
            main_window.select_shell_row("wnd[0]/usr/shellContent/shell", 1)
            ```
        """
        if self.data_present is False:
            raise ValueError("Data was be found in shell table")

        try:
            self._session_handle.findById(self.table_element).currentCellRow = index
            self._session_handle.findById(self.table_element).selectedRows = index

        except Exception as e:
            raise exceptions.ActionException(
                f"Error selecting row with index {index}: {e}"
            )


    def change_checkbox(self, checkbox: str, flag: bool) -> None:
        """
        Sets checkbox in a shell table

        Args:
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
            self._session_handle.findById(self.table_element).changeCheckbox(checkbox, "1", flag)

        except Exception as e:
            raise exceptions.ActionException(f"Error setting element {self.table_element}: {e}")


class ShellTableRowIterator:
    """
    Iterator for shell table rows
    """
    def __init__(self, data: pl.DataFrame) -> None:
        self.data = data
        self.index = 0

    def __iter__(self) -> Self:
        return self

    def __next__(self) -> dict[str, Any]:
        if self.index >= self.data.shape[0]:
            raise StopIteration

        value = self.data.row(self.index, named=True)
        self.index += 1

        return value
