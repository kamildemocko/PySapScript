import os
from time import sleep

from pysapscript.pysapscript import Sapscript
from pysapscript.types_.types import NavigateAction


class TestRuns:
    def __init__(self):
        self.pss = Sapscript()
        pwd = os.getenv("sap_sq8_006_robot01_pwd")

        # Turns on
        self.pss.launch_sap(
            sid="SQ8",
            client="006",
            user="robot01_t",
            password=str(pwd),
            quit_auto=False,
            language="sk"
        )

        self.window = self.pss.attach_window(0, 0)

    def test_runs(self):
        # Basic actions
        self.window.start_transaction("se16")
        self.window.write("wnd[0]/usr/ctxtDATABROWSE-TABLENAME", "LFA1")
        self.window.navigate(NavigateAction.enter)

        # Table
        self.window.write("wnd[0]/usr/txtMAX_SEL", "20")
        self.window.press("wnd[0]/tbar[1]/btn[8]")

        # IF NOT AVL-GRID, switch manually
        table_no_load = self.window.read_shell_table("wnd[0]/usr/cntlGRID1/shellcont/shell", False)
        assert table_no_load.rows > 0
        assert table_no_load.columns > 0

        table = self.window.read_shell_table("wnd[0]/usr/cntlGRID1/shellcont/shell")
        print(f"str: {str(table)}, repr: {repr(table)}")
        print(f"rows: {table.rows}, columns: {table.columns}")
        print(f"polars: {type(table.to_polars_dataframe())}, "
              f"pandas: {type(table.to_pandas_dataframe())}, "
              f"dict: {type(table.to_dict())}")
        print(f"column names: {table.get_column_names()}")

        table.select_rows([1, 3, 5])
        table.select_row(2)
        table.select_all()
        table.clear_selection()

        # export with dropdown
        table.press_context_menu_item("Tabuľková kalkulácia...", item_type="text")
        self.window.set_dropdown(
            "wnd[1]/usr/cmbG_LISTBOX", 
            "10 Excel (vo formáte Office 2007 XLSX)",
            value_type="text",
        )
        self.window.press("wnd[1]/tbar[0]/btn[12]")

        self.window.navigate(NavigateAction.back)
        self.window.navigate(NavigateAction.back)
        self.window.navigate(NavigateAction.back)

        # Non data shell
        self.window.start_transaction("EEDM02")
        nondatashell = self.window.read_shell_table("wnd[0]/usr/subFULLSCREEN_SS:SAPLEEDM_DLG_FRAME:0200/subSUBSCREEN_TREE:SAPLEEDM_TREESELECT:0200/cntlTREE_CONTAINER/shellcont/shell/shellcont[0]/shell")
        nondatashell.press_button("REFRESH_TREE")
        self.window.navigate(NavigateAction.back)

        self.pss.quit()


if __name__ == "__main__":
    TestRuns().test_runs()
