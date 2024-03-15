from pysapscript.pysapscript import Sapscript
from pysapscript.types_.types import NavigateAction


class TestRuns:
    def __init__(self):
        self.pss = Sapscript()
        self.window = self.pss.attach_window(0, 0)

    def run_test(self):
        value = self.window.read_html_viewer(
            "wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[0]/shell"
        )
        print(value)


if __name__ == "__main__":
    TestRuns().run_test()
