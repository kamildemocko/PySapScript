SAP scripting for use in Python.  
Can perform different actions in SAP GUI client on Windows.

[Github - https://github.com/kamildemocko/PySapScript](https://github.com/kamildemocko/PySapScript)

# Installation

## PyPI

```cmd
pip install pysapscript  # pip
uv add pysapscript       # uv
```

## Local

```cmd
git clone https://github.com/kamildemocko/PySapScript
cd PySapScript
uv sync
uv pip install -e .
```

# Documentation

[https://kamildemocko.github.io/PySapScript/](https://kamildemocko.github.io/PySapScript/)

## Local

```cmd
pdoc --html --output-dir docs .\src\pysapscript\
```

# Usage

## Create pysapscript object

```python
import pysapscript

sapscript = pysapscript.Sapscript()
```

parameter `default_window_title: = "SAP Easy Access"`

## Launch Sap

```python
sapscript.launch_sap(
    sid="SQ4",
    client="012",
    user="robot_t",
    password=os.getenv("secret_password")
)
```

additional parameters:

`root_sap_dir = Path(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui")`  
`maximise = True`  
`language = "de"`  
`timeout = 30`  
`quit_auto = True`

## Attach to an already opened window:

```python
from pysapscript.window import Window

window: Window = sapscript.attach_window(0, 0)
```

positional parameters (0, 0) -> (connection, session)

## Quitting SAP:

- pysapscript will automatically quit if not manually specified in `launch_sap` parameter
- manual quitting method: `sapscript.quit()`

## Performing action:

**element**: use SAP path starting with `wnd[0]` for element arguments, for example `wnd[0]/usr/txtMAX_SEL`  
- element paths can be found by recording a sapscript with SAP GUI or by applications like [SAP Script Tracker](https://tracker.stschnell.de/)

```python
window = sapscript.attach.window(0, 0)

window.maximize()
window.restore()
window.close()

window.start_transaction(value)
window.navigate(NavigateAction.enter)
window.navigate(NavigateAction.back)
status = window.read_statusbar()

window.write(element, value)
window.press(element)
window.press_tab([focus_element="wnd[0]", backwards=False])
window.focus(element)
window.send_v_key(value[, focus_element=True, value=0])
window.select(element)
selected = window.is_selected(element)
window.set_checkbox(value)
window.read(element)
window.visualize(element[, seconds=1])
window.exists(element)

window.set_dropdown(element, "02")
window.set_dropdown(element, "Excel File XLSX", value_type="text")

window.show_msgbox(title, message)

table: ShellTable = window.read_shell_table(element)
tree: TreeTable = window.read_shell_tree(element)
html_content = window.read_html_viewer(element)
```

## Table actions

ShellTable uses polars, but can also be return pandas or dictionary

### ShellTable

```python
from pysapscript.shell_table import ShellTable

table: ShellTable = window.read_shell_table()

# shape
table.rows
table.columns

# getters
table.to_dict()
table.to_dicts()
table.to_polars_dataframe()
table.to_pandas_dataframe()

table.cell(row_value, col_value_or_name)
table.get_column_names()

# actions
table.load()
table.press_button(value)
table.click_current_cell()
table.select_rows([0, 1, 2])
table.select_row(1)
table.select_all()
table.clear_selection()
table.change_checkbox(element, value)

table.press_context_menu_item("%XXL")
table.press_context_menu_item("Excel File...", item_type="text")
```

## Tree actions

Holds data in a list of *Node*

### ShellTree

```python
tree: ShellTree = window.read_shell_tree()

# slicing
tree[index]
tree[start:stop:step]

# getters
node = tree.get_node_by_key("         7")
node = tree.get_node_by_label("Name 1")
list_of_nodes = tree.get_nodes()
list_of_node_folders = tree.get_node_folders()
list_of_node_not_foldres = tree.get_node_not_folders()

# other
tree.select_all()
tree.unselect_all()

tree.expand_all()
tree.collapse_all()
```

### Node

```python
node = tree.get_node_by_label("Name 1")

children_nodes = node.get_children()

node.select()
node.unselect()
node.expand()
node.collapse()

node.double_click()
```
