import os
from time import sleep
from pprint import pprint

from pysapscript.pysapscript import Sapscript
from pysapscript.types_.types import NavigateAction
from pysapscript.shell_tree import Node


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

        print("slicing")
        node = tree[5]
        assert isinstance(node, Node)
        assert  node.label =="2. typ výberu"

        nodes_slice = tree[5:7]
        assert isinstance(nodes_slice, list)
        assert len(nodes_slice) == 2

        empty_slice = tree[100:1001]
        assert isinstance(empty_slice, list)
        assert len(empty_slice) == 0

        print("gets")
        nodes = tree.get_nodes()
        assert isinstance(nodes, list)
        assert len(nodes) > 0

        node = tree.get_node_by_label("Hodnota výberu 2")
        assert node is not None
        assert node.key == "          7"

        node = tree.get_node_by_key("          7")
        assert node is not None
        assert node.label == "Hodnota výberu 2"

        node = tree.get_node_by_label("Does not exist")
        assert node is None

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
