Attribute VB_Name = "Module1"
Dim T As String
Dim foundst, table_range As Range
Dim lcol, check_result, found_row, current_row As Long
Sub main()

Application.ScreenUpdating = False

Dim macro_range, input_range, negative_range As Range
Dim lrow, last_row, macro_column_start As Long
Dim count_year As Long

last_row = Sheets("input").Cells(Sheets("input").Rows.Count, 2).End(xlUp).Row

lcol = 20


Do Until lrow = 1
    lrow = Cells(Sheets("input").Rows.Count, lcol).End(xlUp).Row
    lcol = lcol + 1
    Loop
    
lcol = lcol - 2

For current_row = 2 To last_row
    Call add_missing_value
    T = Sheets("input").Cells(current_row, "b").Value
    count_year = WorksheetFunction.CountA(Range(Sheets("input").Cells(current_row, "l"), Sheets("input").Cells(current_row, lcol)))
    Call check_value
    If check_result = 1 And count_year > 3 Then
    Sheets("input").Cells(current_row, "ad").Value = 1
    macro_column_start = lcol + 1
    Set macro_range = get_macro_range(T)
    Set input_range = Range(Sheets("input").Cells(current_row, "l"), Sheets("input").Cells(current_row, lcol))
    For count_year = (lcol + 1) To 28
    
        Sheets("input").Cells(current_row, count_year).Value = WorksheetFunction.Forecast_Linear(Sheets("macro").Cells(found_row, macro_column_start).Value, input_range, macro_range)
        macro_column_start = macro_column_start + 1
        Next count_year
    End If
    Set negative_range = Range(Sheets("input").Cells(current_row, "l"), Sheets("input").Cells(current_row, "ab"))
    Sheets("input").Cells(current_row, "ae").Value = WorksheetFunction.CountIf(negative_range, "<" & 0)
    
    
    Next current_row
    
    Application.ScreenUpdating = True
    
    Set table_range = Range(Sheets("input").Cells(1, 1), Sheets("input").Cells(last_row, "ae"))
    table_range.AutoFilter field:=31, Criteria1:=">0"
    

End Sub



Sub check_outlier()
Application.ScreenUpdating = False

Dim j, check_row, last_check_row, count_year, lcot, ldong As Long
Dim input_range, clear_range As Range
Dim quad1, quad3, iqr, lim1, lim2, mean_value As Double

lcot = 19


Do Until ldong = 1
    ldong = Cells(Sheets("input").Rows.Count, lcot).End(xlUp).Row
    lcot = lcot + 1
    Loop
lcot = lcot - 2

last_check_row = Cells(Sheets("input").Rows.Count, 2).End(xlUp).Row
Set clear_range = Sheets("input").Range(Sheets("input").Cells(2, "a"), Sheets("input").Cells(last_check_row, "ab"))
clear_range.Interior.Color = xlNone


For check_row = 2 To last_check_row
   

Set input_range = Sheets("input").Range(Sheets("input").Cells(check_row, "l"), Sheets("input").Cells(check_row, lcot))
count_year = WorksheetFunction.CountA(input_range)

If count_year >= 3 Then

 
quad1 = WorksheetFunction.Quartile(input_range, 1)
quad3 = WorksheetFunction.Quartile(input_range, 3)
iqr = quad3 - quad2
lim1 = quad1 - (1.5 * iqr)
lim2 = quad3 + (1.5 * iqr)


For j = 12 To lcot
    If Not IsEmpty(Sheets("input").Cells(check_row, j).Value) And (Sheets("input").Cells(check_row, j).Value > lim2 Or Sheets("input").Cells(check_row, j).Value < lim1) Then
    Cells(check_row, j).Interior.ColorIndex = 4
    End If
    Next j
End If

Next check_row

Call count_outlier

Application.ScreenUpdating = True

Set table_range = Range(Sheets("input").Cells(1, 1), Sheets("input").Cells(last_check_row, "ae"))
table_range.AutoFilter field:=29, Criteria1:=">0"

End Sub

Sub count_outlier()

Dim j, check_row, last_check_row, number_of_outlier As Long

last_check_row = Cells(Sheets("input").Rows.Count, 2).End(xlUp).Row

For check_row = 2 To last_check_row
    number_of_outlier = 0
    For j = 12 To 28
    If Cells(check_row, j).Interior.ColorIndex = 4 Then
    number_of_outlier = number_of_outlier + 1
    End If
    Next j
    Sheets("input").Cells(check_row, "ac").Value = number_of_outlier
    
Next check_row

End Sub

Sub add_missing_value()

Dim i As Long
Dim crow As Long
Dim lrow As Long




crow = current_row

For i = 13 To (lcol - 1)
    If IsEmpty(Sheets("input").Cells(crow, i).Value) And Sheets("input").Cells(crow, i - 1).Value > 0 And Sheets("input").Cells(crow, i + 1).Value > 0 Then
    Sheets("input").Cells(crow, i).Value = Sheets("input").Cells(crow, i - 1).Value * ((Sheets("input").Cells(crow, i + 1).Value / Sheets("input").Cells(crow, i - 1).Value) ^ (1 / 2))
    End If
    If IsEmpty(Sheets("input").Cells(crow, i).Value) And IsEmpty(Sheets("input").Cells(crow, i + 1).Value) And Sheets("input").Cells(crow, i + 2).Value > 0 And Sheets("input").Cells(crow, i - 1).Value > 0 Then
    Sheets("input").Cells(crow, i).Value = Sheets("input").Cells(crow, i - 1).Value * ((Sheets("input").Cells(crow, i + 2).Value / Sheets("input").Cells(crow, i - 1).Value) ^ (1 / 3))
    Sheets("input").Cells(crow, i + 1).Value = Sheets("input").Cells(crow, i).Value * ((Sheets("input").Cells(crow, i + 2).Value / Sheets("input").Cells(crow, i - 1).Value) ^ (1 / 3))
    End If
    
    If IsEmpty(Sheets("input").Cells(crow, i).Value) And IsEmpty(Sheets("input").Cells(crow, i + 1).Value) And IsEmpty(Sheets("input").Cells(crow, i + 2).Value) And Sheets("input").Cells(crow, i + 3).Value > 0 And Sheets("input").Cells(crow, i - 1).Value > 0 Then
    Sheets("input").Cells(crow, i).Value = Sheets("input").Cells(crow, i - 1).Value * ((Sheets("input").Cells(crow, i + 3).Value / Sheets("input").Cells(crow, i - 1).Value) ^ (1 / 4))
    Sheets("input").Cells(crow, i + 1).Value = Sheets("input").Cells(crow, i).Value * ((Sheets("input").Cells(crow, i + 3).Value / Sheets("input").Cells(crow, i - 1).Value) ^ (1 / 4))
    Sheets("input").Cells(crow, i + 2).Value = Sheets("input").Cells(crow, i + 1).Value * ((Sheets("input").Cells(crow, i + 3).Value / Sheets("input").Cells(crow, i - 1).Value) ^ (1 / 4))
    
    End If
    
    Next i
    
End Sub

Sub check_value()
Dim lrow2 As Long
Dim range2 As Range


lrow2 = Sheets("macro").Cells(Sheets("macro").Rows.Count, 4).End(xlUp).Row

Set range2 = Sheets("macro").Range(Sheets("macro").Cells(2, 4), Sheets("macro").Cells(lrow2, 4))

Set foundst = range2.Find(T)
If (Not foundst Is Nothing) Then
    check_result = 1
    Else: check_result = 0
    End If

End Sub

Public Function get_macro_range(ByVal ten As String) As Range
Dim range1 As Range
Dim lrow1 As Long
lrow1 = Sheets("macro").Cells(Sheets("macro").Rows.Count, 4).End(xlUp).Row

Set range1 = Sheets("macro").Range(Sheets("macro").Cells(2, 4), Sheets("macro").Cells(lrow1, 4))
found_row = range1.Find(ten).Row
Set get_macro_range = Sheets("macro").Range(Sheets("macro").Cells(found_row, "l"), Sheets("macro").Cells(found_row, lcol))
Exit Function


End Function

Sub count_outlier_again()
Application.ScreenUpdating = False

Dim j, check_row, last_check_row, count_year As Long
Dim input_range As Range
Dim quad1, quad3, iqr, lim1, lim2, mean_value As Double

last_check_row = Cells(Sheets("input").Rows.Count, 2).End(xlUp).Row

For check_row = 2 To last_check_row
   

Set input_range = Sheets("input").Range(Sheets("input").Cells(check_row, "l"), Sheets("input").Cells(check_row, "ab"))
count_year = WorksheetFunction.CountA(input_range)

If count_year >= 3 Then

 
quad1 = WorksheetFunction.Quartile(input_range, 1)
quad3 = WorksheetFunction.Quartile(input_range, 3)
iqr = quad3 - quad2
lim1 = quad1 - (1.5 * iqr)
lim2 = quad3 + (1.5 * iqr)


For j = 12 To 28
    Sheets("input").Cells(check_row, j).Interior.Color = xlNone
    If Not IsEmpty(Sheets("input").Cells(check_row, j).Value) And (Sheets("input").Cells(check_row, j).Value > lim2 Or Sheets("input").Cells(check_row, j).Value < lim1) Then
    Sheets("input").Cells(check_row, j).Interior.ColorIndex = 4
    End If
    Next j
End If

Next check_row

Call count_outlier

Application.ScreenUpdating = True


Set table_range = Range(Sheets("input").Cells(1, 1), Sheets("input").Cells(last_check_row, "ae"))
table_range.AutoFilter field:=29, Criteria1:=">0"


End Sub

Sub fill_rest()
Dim i, lrow, lcot As Long
lcot = lcol

If Sheets("input").AutoFilterMode Then
    Sheets("input").AutoFilterMode = False
    End If
    

lrow = Cells(Sheets("input").Rows.Count, 2).End(xlUp).Row
For i = 2 To lrow
    If IsEmpty(Sheets("input").Cells(i, lcot).Value) And Not IsEmpty(Sheets("input").Cells(i, lcot + 1).Value) And Not IsEmpty(Sheets("input").Cells(i, lcot - 1).Value) Then
    Sheets("input").Cells(i, lcot).Value = Sheets("input").Cells(i, lcot - 1).Value * ((Sheets("input").Cells(i, lcot + 1).Value / Sheets("input").Cells(i, lcot - 1).Value) ^ (1 / 2))
    End If
    
    If IsEmpty(Sheets("input").Cells(i, lcot).Value) And IsEmpty(Sheets("input").Cells(i, lcot - 1).Value) And Not IsEmpty(Sheets("input").Cells(i, lcot + 1).Value) And Not IsEmpty(Sheets("input").Cells(i, lcot - 2).Value) Then
    Sheets("input").Cells(i, lcot - 1).Value = Sheets("input").Cells(i, lcot - 2).Value * ((Sheets("input").Cells(i, lcot + 1).Value / Sheets("input").Cells(i, lcot - 2).Value) ^ (1 / 3))
    Sheets("input").Cells(i, lcot).Value = Sheets("input").Cells(i, lcot - 1).Value * ((Sheets("input").Cells(i, lcot + 1).Value / Sheets("input").Cells(i, lcot - 2).Value) ^ (1 / 3))
    End If
    
    If IsEmpty(Sheets("input").Cells(i, lcot).Value) And IsEmpty(Sheets("input").Cells(i, lcot - 1).Value) And IsEmpty(Sheets("input").Cells(i, lcot - 2).Value) And Not IsEmpty(Sheets("input").Cells(i, lcot + 1).Value) And Not IsEmpty(Sheets("input").Cells(i, lcot - 3).Value) Then
    Sheets("input").Cells(i, lcot - 2).Value = Sheets("input").Cells(i, lcot - 3).Value * ((Sheets("input").Cells(i, lcot + 1).Value / Sheets("input").Cells(i, lcot - 3).Value) ^ (1 / 4))
    Sheets("input").Cells(i, lcot - 1).Value = Sheets("input").Cells(i, lcot - 2).Value * ((Sheets("input").Cells(i, lcot + 1).Value / Sheets("input").Cells(i, lcot - 3).Value) ^ (1 / 4))
    Sheets("input").Cells(i, lcot).Value = Sheets("input").Cells(i, lcot - 1).Value * ((Sheets("input").Cells(i, lcot + 1).Value / Sheets("input").Cells(i, lcot - 3).Value) ^ (1 / 4))
    End If
    Next i


End Sub

Sub remove_negative()
Dim i, lrow, lcot As Long
Dim range2 As Range

Application.ScreenUpdating = False

lcot = lcol + 1


lrow = Cells(Sheets("input").Rows.Count, 2).End(xlUp).Row
For i = 2 To lrow
    If Sheets("input").Cells(i, "ae").Value > 0 Then
    Set range2 = Sheets("input").Range(Sheets("input").Cells(i, lcot), Sheets("input").Cells(i, "ab"))
    range2.ClearContents
    Sheets("input").Cells(i, "ae").Value = 0
    End If
    Next i


    
Application.ScreenUpdating = True

End Sub

Sub Auto_Open()

form1.Show


End Sub

