# ExcelDarkMode
Customisable Dark Mode for Excel cell backgrounds. 

VBA based but will work on non-macro workbooks if you save it to your Personal workbook (see below).

This works by replacing the colour definitions for styles already defined in the workbook (rather than creating new styles). 

Don't forget to toggle a workbook back to light mode before you share it with uninitiated friends!

![ExcelDarkMode](https://user-images.githubusercontent.com/17323155/122040068-4c7b3780-cdcf-11eb-8d70-ba46c97d1f05.gif)

# Setup

1. Copy the code from [DarkMode.vba](https://raw.githubusercontent.com/stu0292/ExcelDarkMode/main/DarkMode.vba) to a new module in your [PERSONAL.XLSB](https://support.microsoft.com/en-gb/office/create-and-save-all-your-macros-in-a-single-workbook-66c97ab3-11c2-44db-b021-ae005a9bc790), save. 
2. Run the macro ToggleDarkMode to switch between dark and light modes
3. For ease of use, add the macro to your [Quick Access bar](https://www.howtogeek.com/232620/how-to-add-a-macro-to-the-quick-access-toolbar-in-office/)

# Customisation

To change the dark mode colours, change the Hex colour code parameters in the `DarkMode` Function:

```
   Call ApplyDarkStyle(styleName:="Normal", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF", borderColorHex:="#454545")
```

To find the Hex codes for your favourite colours, use the font colour 'More Colors' picker (and copy the Hex code from the bottom):

![image](https://user-images.githubusercontent.com/17323155/122032272-60bb3680-cdc7-11eb-858d-32e3b2fedf65.png)

To set dark mode for a new style, copy the `Call ApplyDarkStyle...` line and provide the style name. Don't forget to add the style to the `LightMode` Function as well. 


# What if it doesn't work for all cells?

This approach only works for [cell styles](https://support.microsoft.com/en-us/office/apply-create-or-remove-a-cell-style-472213bf-66bd-40c8-815c-594f0f90cd22). 
Cells which have had font/fill colours applied *in addition* to the cell style will not be darkened, since the additional formatting overrides cell style formatting.  

For dark mode to work for these cells, you need to re-apply a cell style to these cells. Beware that this will remove ALL additional formatting on the cell that isn't defined in a style.

There's a macro `ResetStyles` to assist removing additional formatting so you can apply styles properly to a pre-existing workbook. Use with caution. Either save a copy of the file or copy the formatting/sheet to a new workbook to save the original formatting until you're happy with your changes. In some cases you may be better off selecting individual regions, clearing formatting manually.

It may take a bit of time to undo formatting and replace with proper cell styles for a pre-existing Workbook but this is good practice and should hopefully be a one-off activity. Or just accept that not all cells will be darkened. 

## Font colour isn't lightened

If the cell background gets darkedned but the font colour doesn't change (ie dark text on dark background), it's most likely because the font colour has been changed for those cells. This is a variation of the case in the section above - to fix you'd need to reapply a style to these cells  

# Feedback welcome

To send feedback, head to the GitHub repository and create a new issue. 

