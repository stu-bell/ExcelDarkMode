' https://github.com/stu-bell/ExcelDarkMode
' Copyright (c) 2021 Stuart Bell
' Licenced under the MIT licence: https://github.com/stu-bell/ExcelDarkMode/blob/main/LICENSE

' Inverts the colors of styles. Work in progress!

Sub ToggleDarkMode2()

    Call InvertStyleColors
    
    'FIXME
    'Call invertTableStyleColors
End Sub

' invert colors of table styles
' FIXME some table styles elements seem to have no fill, ie they take a transparent background, which makes it unreadable given the way it currently interacts with the inversion of the regular styles
Sub invertTableStyleColors()

    ' Flip all tables to dark equivalent
    Dim tbl As ListObject
    Dim sht As Worksheet
    Dim styleName As String
      For Each sht In ActiveWorkbook.Worksheets
        For Each tbl In sht.ListObjects
            tbl.TableStyle = GetInvertTableStyleName(tbl.TableStyle)
        Next tbl
      Next sht

'   FIXME: trying to invert colors of individual table style elements but I can't get it to work
'    For Each st In ActiveWorkbook.TableStyles
'        Dim i As Integer
'        Dim el As TableStyleElement
'        For i = 1 To st.TableStyleElements.Count
'            If st.TableStyleElements(i).HasFormat = True Then
'                Set el = st.TableStyleElements(i)
'                st.TableStyleElements(i).Interior.color = InvertColor(st.TableStyleElements(i).Interior.color)
'            End If
'        Next i
'    Next st

End Sub

' quick and dirty way to try and match a table theme with a corresponding dark theme
Function GetInvertTableStyleName(tableStyleName As String)

If InStr(tableStyleName, "Dark") Then
    GetInvertTableStyleName = Replace(tableStyleName, "Dark", "Light")
ElseIf InStr(tableStyleName, "Light") Then
    GetInvertTableStyleName = Replace(tableStyleName, "Light", "Dark")
End If

End Function


' Inverts font and interior colors for each style in the workbook
Sub InvertStyleColors()

For Each sty In ActiveWorkbook.Styles

    ' todo include border color properties from DarkMode 1

    ' Todo: can we find a smarter way of choosing the dark/light equivalent colors than just simple inverts?
    sty.Font.color = InvertColor(sty.Font.color)
    
    ' not all styles have a pattern included by default (eg headings) so we include them here
    ' not sure if this will mess up other formatting though
    sty.IncludePatterns = True
    sty.Interior.color = InvertColor(sty.Interior.color)
    
Next sty

End Sub

' Inverts a color code
Function InvertColor(color As Long)
     r = color And 255
     g = color \ 256 And 255
     b = color \ 256 ^ 2 And 255
        
    InvertColor = RGB(255 - r, 255 - g, 255 - b)
End Function

Sub InvertColor_test()
    MsgBox _
        InvertColor(RGB(200, 100, 50)) = RGB(55, 155, 205) And _
        InvertColor(RGB(200, 100, 50)) = RGB(55, 155, 205) And _
        InvertColor(RGB(0, 0, 0)) = RGB(255, 255, 255)
End Sub


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



