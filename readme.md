# Description
SAP scripting for Python automatization

# Installation
```pip
pip install pysapscript
```

# Usage
## Create pysapscript object
```python
pss = pysapscript.Sapscript()
```
parameter `default_window_title: = "SAP Easy Access"`

## Launch Sap
```python
pss.launch_sap(
    sid="SQ4",
    client="012",
    user="robot_t",
    password=os.getenv("secret_password")
)
```
additional parameters:
```python
root_sap_dir = Path(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui")
maximise = True
quit_auto = True
```

## Attach to window:
```python
window = pss.attach_window(0, 0)
```
positional parameters (0, 0) -> (connection, session)

## Quitting SAP:
- will automatically quit if not specified differently
- manual quitting: `pss.quit()`

## Performing action:
```
window.write(element, value)
window.press(element)
window.select(element)
window.read(element)
window.read_shell_table(element)
window.press_shell_button(element, button_name)
window.change_shell_checkbox(element, checkbox_name, boolean)
```
- use SAP path `wnd[0].....` for elements

And another available actions...
- close window, open new window, start transaction, navigate, maximize
    
