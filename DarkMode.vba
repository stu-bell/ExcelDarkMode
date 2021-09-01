' Customisable dark mode for Excel cells
' https://github.com/stu0292/ExcelDarkMode
' Copyright (c) 2021 Stuart Bell
' Licenced under the MIT licence: https://github.com/stu0292/ExcelDarkMode/blob/main/LICENSE
' Only modifies cell styles. Will not change colors of cells that have been formatted separately.
' To include custom formatted cells in dark mode, create a new style for that formatting and include the style in this module
' Color codes for each style must be inserted into the code below  for both DarkMode and LightMode (see comments in Functions below)
' Original table styles are not preserved when switching back to light mode - you'll need to specify the default light style in code or use sub SetWorkbookTableStyle
' Save this macro in your PERSONAL.XLSB (and add it to your quick access bar!) so you can use dark mode in any new workbook, including non-macro enabled ones
Function DarkMode()
    ' Set all tables to this dark table style
    Call SetAllTableStyle("TableStyleDark2")

    ' Dark colors for each style
    Call UpdateStyleColors(styleName:="Normal", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF", borderColorHex:="#454545")
    Call UpdateStyleColors(styleName:="Heading 1", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF")
    Call UpdateStyleColors(styleName:="Heading 2", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF")
    Call UpdateStyleColors(styleName:="Heading 3", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF")
    Call UpdateStyleColors(styleName:="Heading 4", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF")
    Call UpdateStyleColors(styleName:="Title", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF")
    Call UpdateStyleColors(styleName:="Total", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF")
    Call UpdateStyleColors(styleName:="Note", fillColorHex:="#B2B2B2", fontColorHex:="#000000", borderColorHex:="#454545")
    Call UpdateStyleColors(styleName:="Explanatory Text", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF", borderColorHex:="#454545")
End Function

Function LightMode()
    ' Set all tables to this light table style
    Call SetAllTableStyle("TableStyleMedium9")

    ' Light colors for each style
    Call UpdateStyleColors(styleName:="Normal", fillColorHex:="#FFFFFF", fontColorHex:="#000000", borderLineStyle:=xlNone, interiorPattern:=xlNone)
    Call UpdateStyleColors(styleName:="Heading 1", fontColorHex:="#44546A", interiorPattern:=xlNone)
    Call UpdateStyleColors(styleName:="Heading 2", fontColorHex:="#44546A", interiorPattern:=xlNone)
    Call UpdateStyleColors(styleName:="Heading 3", fontColorHex:="#44546A", interiorPattern:=xlNone)
    Call UpdateStyleColors(styleName:="Heading 4", fontColorHex:="#44546A", interiorPattern:=xlNone)
    Call UpdateStyleColors(styleName:="Title", fontColorHex:="#44546A", interiorPattern:=xlNone)
    Call UpdateStyleColors(styleName:="Total", fontColorHex:="#000000", interiorPattern:=xlNone)
    Call UpdateStyleColors(styleName:="Note", fillColorHex:="#FFFFCC", fontColorHex:="#000000", borderColorHex:="#B2B2B2")
    Call UpdateStyleColors(styleName:="Explanatory Text", fontColorHex:="#7F7F7F", borderColorHex:="#454545", interiorPattern:=xlNone)


End Function

Function DarkModeWithBackup()
    
    ' Set all tables to this dark table style
    Call SetAllTableStyle("TableStyleDark2")

    ' List calls to dark styles
    Call ApplyDarkStyle(styleName:="Normal", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF", borderColorHex:="#454545")
    Call ApplyDarkStyle(styleName:="Heading 1", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF")
    Call ApplyDarkStyle(styleName:="Heading 2", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF")
    Call ApplyDarkStyle(styleName:="Heading 3", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF")
    Call ApplyDarkStyle(styleName:="Heading 4", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF")
    Call ApplyDarkStyle(styleName:="Title", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF")
    Call ApplyDarkStyle(styleName:="Total", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF")
    Call ApplyDarkStyle(styleName:="Note", fillColorHex:="#B2B2B2", fontColorHex:="#000000", borderColorHex:="#454545")
    Call ApplyDarkStyle(styleName:="Explanatory Text", fillColorHex:="#2E3440", fontColorHex:="#FFFFFF", borderColorHex:="#454545")
    
    ' This should fail without error, as the style doesn't exist
    'Call ApplyDarkStyle(styleName:="noexistlsakjalsdfkj", fillColorHex:="#000000")
End Function
Function LightModeFromBackup()
    
    ' Set all tables to this light table style
    Call SetAllTableStyle("TableStyleLight1")

    ' List calls to light style FIXME: loop this
    Call RestoreLightStyle("Normal")
    Call RestoreLightStyle("Heading 1")
    Call RestoreLightStyle("Heading 2")
    Call RestoreLightStyle("Heading 3")
    Call RestoreLightStyle("Heading 4")
    Call RestoreLightStyle("Title")
    Call RestoreLightStyle("Total")
    Call RestoreLightStyle("Note")
    Call RestoreLightStyle("Explanatory Text")
End Function

Sub ToggleDarkMode()
    
    Application.ScreenUpdating = False
    
    ' Create a custom property to save state of Dark/Light mode in the workbook
    Dim flag As String
    flag = "DARK_MODE_0292"
    If Not CustomPropertyExists(flag) Then
        ActiveWorkbook.CustomDocumentProperties.Add Name:=flag, Value:=0, _
            LinkToContent:=False, _
            Type:=msoPropertyTypeNumber
    End If
    
    ' Toggle state based on flag
    If ActiveWorkbook.CustomDocumentProperties(flag).Value = 1 Then
        ' Dark to Light
        ActiveWorkbook.CustomDocumentProperties(flag).Value = 0
        Call LightMode
    Else
        ' Light to Dark
        ActiveWorkbook.CustomDocumentProperties(flag).Value = 1
        Call DarkMode
    End If

    Application.ScreenUpdating = True

End Sub

'Resets all tables to the style named here
Sub SetWorkbookTableStyle()
    Dim tabStyleName As String
    tabStyleName = InputBox("This table style will be applied to all tables in the workbook." & vbCrLf & "Available table style names can be found in the Table Design ribbon (just remove spaces from the name in this box)", "Table Style Name", "TableStyleLight1")
    Call SetAllTableStyle(tabStyleName)
End Sub

' Resets formatting of cells to their original style (resets all formatting done on top of ANY style)
' If the workbook hasn't had styles properly used you'll loose a lot of formatting
' Use with caution!
Sub ResetStyles()
' https://jkp-ads.com/Articles/styles06.asp
    Dim oCell As Range
    Dim oSh As Worksheet
    If MsgBox("This will erase all additional formatting on top of the existing cell styles in the selected sheets." & vbNewLine & _
              "If you're not sure, Cancel this and save a copy of your workbook", _
              vbCritical + vbOKCancel + vbDefaultButton2, "This step is not reversible") = vbOK Then
    Application.ScreenUpdating = False
        For Each oSh In ActiveWindow.SelectedSheets
            For Each oCell In oSh.UsedRange.Cells
                If oCell.MergeArea.Cells.Count = 1 Then
                    ' reapply original style and remove additional formatting
                    oCell.Style = CStr(oCell.Style)
                End If
            Next
        Next
    End If
    Application.ScreenUpdating = True
End Sub



' Change the color properties of the style
' To modify a new property (eg font name) set the property as a new optional arg
' All style params must be optional and tested for with `If Not IsMissing(paramName)`
Function UpdateStyleColors(styleName As String, _
    Optional fillColorHex As String, _
    Optional fontColorHex As String, _
    Optional borderColorHex As String, _
    Optional borderLineStyle As XlLineStyle, _
    Optional interiorPattern As XlPattern)
    ' Skip styles we haven't defined
    On Error Resume Next
    
    ' Make sure the style includes all of the elements we want to change (eg Heading 1 doesn't include Patterns by default
    With ActiveWorkbook.Styles(styleName)
        .IncludeFont = True
        .IncludeBorder = True
        .IncludePatterns = True
    End With
    
    ' Set the properties of the target style (only if a parameter has been passed for that property)
    ' FIXME can we choose properties dynamically in VBA?
    With ActiveWorkbook.Styles(styleName)
        If Not IsMissing(fillColorHex) Then
            .Interior.Color = HexToRGB(fillColorHex)
        End If
        If Not IsMissing(fontColorHex) Then
            .Font.Color = HexToRGB(fontColorHex)
        End If
        If Not IsMissing(borderColorHex) Then
            .Borders(xlLeft).Color = HexToRGB(borderColorHex)
            .Borders(xlRight).Color = HexToRGB(borderColorHex)
            .Borders(xlBottom).Color = HexToRGB(borderColorHex)
            .Borders(xlTop).Color = HexToRGB(borderColorHex)
        End If
        
        If borderLineStyle <> 0 Then
            .Borders(xlLeft).LineStyle = borderLineStyle
            .Borders(xlRight).LineStyle = borderLineStyle
            .Borders(xlBottom).LineStyle = borderLineStyle
            .Borders(xlTop).LineStyle = borderLineStyle
        End If
        
        If interiorPattern <> 0 Then
            .Interior.Pattern = interiorPattern
        End If
    End With

End Function

' Change the color properties of the style to make them dark. Stores original style colors in a backup style
' To modify a new property (eg font name) set the property as a new optional arg and make sure to add the property definition to backup style (this function), the actual style (this function) and the function RestoreLightStyle
' All style params must be optional and tested for with `If Not IsMissing(paramName)`
Function ApplyDarkStyle(styleName As String, _
    Optional fillColorHex As String, _
    Optional fontColorHex As String, _
    Optional borderColorHex As String)
    ' Skip styles we haven't defined
    On Error Resume Next
    
    ' Make sure the style includes all of the elements we want to change (eg Heading 1 doesn't include Patterns by default
    With ActiveWorkbook.Styles(styleName)
        .IncludeFont = True
        .IncludeBorder = True
        .IncludePatterns = True
    End With
    
    ' Create a backup style for the style, saving the original
    With ActiveWorkbook.Styles.Add(BackupStyleName(styleName))
     If Not IsMissing(fillColorHex) Then
         .Interior.Color = ActiveWorkbook.Styles(styleName).Interior.Color
    End If
     If Not IsMissing(fontColorHex) Then
         .Font.Color = ActiveWorkbook.Styles(styleName).Font.Color
    End If
     If Not IsMissing(borderColorHex) Then
         .Borders(xlLeft).Color = ActiveWorkbook.Styles(styleName).Borders(xlLeft).Color
         .Borders(xlRight).Color = ActiveWorkbook.Styles(styleName).Borders(xlRight).Color
         .Borders(xlBottom).Color = ActiveWorkbook.Styles(styleName).Borders(xlBottom).Color
         .Borders(xlTop).Color = ActiveWorkbook.Styles(styleName).Borders(xlTop).Color
         .Borders(xlLeft).LineStyle = ActiveWorkbook.Styles(styleName).Borders(xlLeft).LineStyle
         .Borders(xlRight).LineStyle = ActiveWorkbook.Styles(styleName).Borders(xlRight).LineStyle
         .Borders(xlBottom).LineStyle = ActiveWorkbook.Styles(styleName).Borders(xlBottom).LineStyle
         .Borders(xlTop).LineStyle = ActiveWorkbook.Styles(styleName).Borders(xlTop).LineStyle
    End If
    End With
    
    ' Backup the interior pattern
    ActiveWorkbook.Styles(BackupStyleName(styleName)).Interior.Pattern = ActiveWorkbook.Styles(styleName).Interior.Pattern
    
    ' Set the properties of the target style (only if a parameter has been passed for that property)
    ' FIXME can we choose properties dynamically in VBA?
    With ActiveWorkbook.Styles(styleName)
        If Not IsMissing(fillColorHex) Then
            .Interior.Color = HexToRGB(fillColorHex)
        End If
        If Not IsMissing(fontColorHex) Then
            .Font.Color = HexToRGB(fontColorHex)
        End If
        If Not IsMissing(borderColorHex) Then
            .Borders(xlLeft).Color = HexToRGB(borderColorHex)
            .Borders(xlRight).Color = HexToRGB(borderColorHex)
            .Borders(xlBottom).Color = HexToRGB(borderColorHex)
            .Borders(xlTop).Color = HexToRGB(borderColorHex)
        End If
    End With

End Function

' Reset each property to the light style from the backup
Function RestoreLightStyle(styleName As String)

    ' Skip styles we haven't defined
    On Error Resume Next

    With ActiveWorkbook.Styles(styleName)
        .Interior.Color = ActiveWorkbook.Styles(BackupStyleName(styleName)).Interior.Color
        .Font.Color = ActiveWorkbook.Styles(BackupStyleName(styleName)).Font.Color
        .Borders(xlLeft).Color = ActiveWorkbook.Styles(BackupStyleName(styleName)).Borders(xlLeft).Color
        .Borders(xlRight).Color = ActiveWorkbook.Styles(BackupStyleName(styleName)).Borders(xlRight).Color
        .Borders(xlBottom).Color = ActiveWorkbook.Styles(BackupStyleName(styleName)).Borders(xlBottom).Color
        .Borders(xlTop).Color = ActiveWorkbook.Styles(BackupStyleName(styleName)).Borders(xlTop).Color
        .Borders(xlLeft).LineStyle = ActiveWorkbook.Styles(BackupStyleName(styleName)).Borders(xlLeft).LineStyle
        .Borders(xlRight).LineStyle = ActiveWorkbook.Styles(BackupStyleName(styleName)).Borders(xlRight).LineStyle
        .Borders(xlBottom).LineStyle = ActiveWorkbook.Styles(BackupStyleName(styleName)).Borders(xlBottom).LineStyle
        .Borders(xlTop).LineStyle = ActiveWorkbook.Styles(BackupStyleName(styleName)).Borders(xlTop).LineStyle
        .Interior.Pattern = ActiveWorkbook.Styles(BackupStyleName(styleName)).Interior.Pattern
    End With

    ' Clean up the backup style
    ActiveWorkbook.Styles(BackupStyleName(styleName)).Delete
End Function

' Wrapper to get the backup name for a known style name
Function BackupStyleName(styleName As String) As String
    BackupStyleName = styleName & "_DARKMODE_BACKUP"
End Function

' Loops through each table in workbook and applies the named table style. Slow for many tables!
Function SetAllTableStyle(styleName As String)
Dim tbl As ListObject
Dim sht As Worksheet
  For Each sht In ActiveWorkbook.Worksheets
    For Each tbl In sht.ListObjects
        tbl.TableStyle = styleName
    Next tbl
  Next sht
End Function

'Convert Hex color codes to RGB for setting color properties in VBA
Function HexToRGB(hex As String)
    
    ' remove optional leading #
    nohash = Replace(hex, "#", "")
    
    ' split hex code into rgb parts
    red = CLng("&H" & Left(nohash, 2))
    green = CLng("&H" & Mid(nohash, 3, 2))
    blue = CLng("&H" & Right(nohash, 2))
    
    HexToRGB = RGB(red, green, blue)
End Function
Sub HexToRGB_test()
    MsgBox _
        RGB(200, 100, 50) = HexToRGB("#C86432") And _
        RGB(200, 100, 50) = HexToRGB("C86432") And _
        RGB(255, 255, 255) = HexToRGB("FFFFFF") And _
        RGB(0, 0, 0) = HexToRGB("000000")
End Sub

Function CustomPropertyExists(propName As String) As Boolean
    Dim wb As Workbook
    Dim docProp As DocumentProperty
    Dim propExists As Boolean
    Set wb = Application.ActiveWorkbook
    For Each docProp In wb.CustomDocumentProperties
        If docProp.Name = propName Then
            propExists = True
            Exit For
        End If
    Next
    CustomPropertyExists = propExists
End Function







