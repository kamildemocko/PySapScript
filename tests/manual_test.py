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
            quit_auto=False
        )

        self.window = self.pss.attach_window(0, 0)

    def test_runs(self):
        # Maximize and Restore
        self.window.restore()
        sleep(1)
        self.window.maximize()
        sleep(1)

        # Basic actions
        self.window.start_transaction("se16")
        self.window.write("wnd[0]/usr/ctxtDATABROWSE-TABLENAME", "LFA1")
        self.window.visualize("wnd[0]/usr/ctxtDATABROWSE-TABLENAME", 2)
        self.window.navigate(NavigateAction.enter)

        self.window.press("wnd[0]/tbar[1]/btn[31]")
        value = self.window.read("wnd[1]/usr/txtG_DBCOUNT")
        print("Element value: " + value)
        self.window.press("wnd[1]/tbar[0]/btn[0]")

        # Table
        self.window.write("wnd[0]/usr/txtMAX_SEL", "20")
        self.window.press("wnd[0]/tbar[1]/btn[8]")
        table = self.window.read_shell_table("wnd[0]/usr/cntlGRID1/shellcont/shell")
        print(table.head())

        self.window.select_shell_rows("wnd[0]/usr/cntlGRID1/shellcont/shell", [1, 3, 5])

        self.window.navigate(NavigateAction.back)
        self.window.navigate(NavigateAction.back)
        self.window.navigate(NavigateAction.back)

        # New window
        self.pss.open_new_window(window_to_handle_opening=self.window)
        sleep(2)  # ! This will not await the window as other window is opened with same title
        window2 = self.pss.attach_window(0, 1)
        window2.start_transaction("SQVI")
        window2.close_window()

        # Quitting
        self.pss.quit()


if __name__ == "__main__":
    TestRuns().test_runs()
