import os
from time import sleep
from pprint import pprint

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
        self.window.start_transaction("fpcpl")

        tree = self.window.read_shell_tree("wnd[0]/shellcont/shellcont/shell/shellcont[0]/shell")
        pprint(f"str: {tree}")
        pprint(f"index 5: {tree[5]}")
        pprint(f"slice 5:7 {tree[5:7]}")
        print(f"len: {len(tree)}")

        print("collapse and expland folders")
        tree.collapse_all()
        sleep(1)
        tree.expand_all()

        print("select and unselect all")
        tree.select_all()
        sleep(1)
        tree.unselect_all()

        self.window.navigate(NavigateAction.back)
        self.pss.quit()


if __name__ == "__main__":
    TestRuns().test_runs()
