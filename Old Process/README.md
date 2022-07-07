# StartFormFilling
Make a GUI to be able to add lots based on

## Concept

- Select file you want to add entries to.
- From said file, generate a drop down of entries based for SLABS and FW Plans
- Have Input forms for the entries
- Drop down menu from Plans and Options

## TODO

- Fix being able to add options to the drop down menu
- Change the first screen to include
    - Number of Lots
    - Sub Job Code
    - Super visor
    - Date
    - Type of pour
- Pull Contract data and auto update the add for

## Known Bugs

- When you add a new option, the drop down menu doesn't update
    - Not added to list, try to make a call to the ExcelProcessor Class and add the options in there.
- When you select an option, no way to remove it. If you delete it, you will get an out of bounds error. Add empty first entry.
- When you add a new option, the drop down menu doesn't update
    - Have each list start with an empty list.



