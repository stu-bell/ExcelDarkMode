# ExcelDarkMode
Customisable Dark Mode for Excel cell backgrounds. 

VBA based but will work on non-macro workbooks if you save it to your Personal workbook (see below).

This works by replacing the colour definitions for styles already defined in the workbook (rather than creating new styles). 

Don't forget to toggle a workbook back to light mode before you share it with uninitiated friends!

# Setup

1. Copy the code from [DarkMode.vba](https://raw.githubusercontent.com/stu0292/ExcelDarkMode/main/DarkMode.vba) to a new module in your [PERSONAL.XLSB](https://support.microsoft.com/en-gb/office/create-and-save-all-your-macros-in-a-single-workbook-66c97ab3-11c2-44db-b021-ae005a9bc790), save. 
2. Run the macro ToggleDarkMode to switch between dark and light modes

# Notes

This only works for [cell styles](https://support.microsoft.com/en-us/office/apply-create-or-remove-a-cell-style-472213bf-66bd-40c8-815c-594f0f90cd22). 
Cells which have had custom formatting applied (custom font colours, custom background fill etc) will not be affected by default. For dark mode to work for these cells, 
remove the custom font and fill colours and create a new custom style for those cells (this is good practice anyway). 

# Customisation

To change the dark mode colours, change the Hex colour code parameters in the `DarkMode` Function:

```
   Call ApplyDarkStyle(styleName:="Normal", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF", borderColorHex:="#454545")
```

To find the Hex codes for your favourite colours, use the font colour 'More Colors' picker (and copy the Hex code from the bottom):

![image](https://user-images.githubusercontent.com/17323155/122032272-60bb3680-cdc7-11eb-858d-32e3b2fedf65.png)

To set dark mode for a new style, copy the `Call ApplyDarkStyle...` line and provide the style name. Don't forget to add the style to the `LightMode` Function as well. 

