# Description
SAP scripting for python


# Usage
## Attaching:
```
sap = Sap()
main_window = sap.attach(0, 0)
```
- (0, 0) -> (connection, session)

## Starting new SAP:
```
sap = Sap()
main_window = sap.start_sap(sap_dir, SID, client, name, password)
````

## Ending session / connection:
```
sap.detach(main_window)
```

## Quitting SAP:
- will automatically quit

## Performing action:
```
sap.element_write(obj1, el_input_transaction, "se16")
sap.element_click(el_enter_button)
sap.element_select(el_enter_button)
```
among another available actions are:
- reading elements
- reading table
- clicking table buttons

also
- opening new window
- starting transaction
    