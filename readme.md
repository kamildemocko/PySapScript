[Github - https://github.com/kamildemocko/PySapScript](https://github.com/kamildemocko/PySapScript)

SAP scripting for use in Python.  
Can perform different actions in SAP GUI client on Windows.


# Documentation

[https://kamildemocko.github.io/PySapScript/](https://kamildemocko.github.io/PySapScript/)

```cmd
pdoc --html --output-dir docs .\src\pysapscript\
```

# Installation

```cmd
uv sync
uv pip install -e .
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

window.write(element, value)
window.press(element)
window.send_v_key(value[, focus_element=True, value=0])
window.select(element)
selected = window.is_selected(element)
window.read(element)
window.set_checkbox(value)
window.visualize(element[, seconds=1])
window.exists(element)

table: ShellTable = window.read_shell_table(element)
html_content = window.read_html_viewer(element)
```

## Table actions

ShellTable uses polars, but can also be return pandas or dictionary

```python
from pysapscript.shell_table import ShellTable

table: ShellTable = window.read_shell_table()

table.rows
table.columns

table.to_dict()
table.to_dicts()
table.to_polars_dataframe()
table.to_pandas_dataframe()

table.cell(row_value, col_value_or_name)
table.get_column_names()

table.load()
table.press_button(value)
table.select_rows([0, 1, 2])
table.select_row(1)
table.change_checkbox(element, value)
```