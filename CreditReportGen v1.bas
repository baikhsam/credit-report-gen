Attribute VB_Name = "Module1"
Sub CreditReportGenerator()
'HongSeok (Sam) Baik for Global Energy Trading Ltd
'7/30/2018
'Macro created to generate credit summary report from a given data excel file

Application.ScreenUpdating = False
Dim ws As Worksheet
Dim data As Worksheet
Set data = Worksheets("DATA")
Dim acc As String
acc = InputBox("For what account do you wish to create a credit report for?", "Input Account Name", "TOTAL MARINE")
acc = UCase(acc)
Dim icredit As Long
icredit = InputBox("What is the initial credit line of " & acc & " ?", "Input Initial Credit", 5000000)
Dim ucredit As Long
ucredit = 0
Set ws = Sheets.Add
Dim today As Variant
today = "1-April-18"

'sorting data by Purchases, Due Date, Account Name
With data.Sort
    .SortFields.Clear
    .SortFields.Add Key:=Range("B1"), Order:=xlAscending
    .SortFields.Add Key:=Range("U1"), Order:=xlAscending
    .SortFields.Add Key:=Range("F1"), Order:=xlAscending
    .SetRange Range("A1:AN9000")
    .Header = xlYes
    .Apply
End With
'creating report skeleton
ws.Name = "Credit Report " & ws.Name
ws.Range("A1") = "Credit Report for " & acc
ws.Range("A1") = UCase(ws.Range("A1"))
ws.Range("A1").Font.Size = 15
ws.Range("A1").Font.Bold = True
ws.Range("A1").Font.Name = "Garamond"
With ws.Columns("A")
.ColumnWidth = .ColumnWidth * 1.6
End With
With ws.Columns("B")
.ColumnWidth = .ColumnWidth * 3
End With
With ws.Columns("C")
.ColumnWidth = .ColumnWidth * 1.6
End With
With ws.Columns("D")
.ColumnWidth = .ColumnWidth * 1.6
End With
With ws.Columns("E")
.ColumnWidth = .ColumnWidth * 1.6
End With
With ws.Columns("F")
.ColumnWidth = .ColumnWidth * 1.6
End With
With ws.Columns("G")
.ColumnWidth = .ColumnWidth * 1.6
End With
With ws.Columns("H")
.ColumnWidth = .ColumnWidth * 1.6
End With
'creating credit table
ws.Range("A3") = acc & " Credit Summary:"
ws.Range("G3") = "Initial Credit Line:"
ws.Range("A4") = "Credit used:"
ws.Range("D4") = "Credit available:"
ws.Range("A7") = "Upcoming Transactions beginning from " & today
ws.Range("A8") = "TRAN DATE:"
ws.Range("B8") = "BARGE:"
ws.Range("C8") = "GRADE:"
ws.Range("D8") = "QTY:"
ws.Range("E8") = "PRICE:"
ws.Range("F8") = "AMT:"
ws.Range("G8") = "TOTAL AMT:"
ws.Range("H8") = "DUE DATE:"
ws.Range("G5") = icredit
'formatting
ws.Range("A3:H5").Borders.LineStyle = xlContinuous
ws.Range("A7:H500").Borders.LineStyle = xlContinuous
ws.Range("A3").Font.Bold = True
ws.Range("G3").Font.Bold = True
ws.Range("A7").Font.Size = 13
ws.Range("A8:H8").Font.Bold = True
ws.Range("G8:H8").Interior.ColorIndex = 8
ws.Range("A4").Interior.ColorIndex = 6
ws.Range("A5").Interior.ColorIndex = 6
ws.Range("D4").Interior.ColorIndex = 4
ws.Range("D5").Interior.ColorIndex = 4
ws.Range("A3").HorizontalAlignment = xlCenter
ws.Range("G3").HorizontalAlignment = xlCenter
ws.Range("A4").HorizontalAlignment = xlCenter
ws.Range("A5").HorizontalAlignment = xlCenter
ws.Range("D4").HorizontalAlignment = xlCenter
ws.Range("D5").HorizontalAlignment = xlCenter
ws.Range("A7").HorizontalAlignment = xlCenter
ws.Range("G5").HorizontalAlignment = xlCenter
ws.Range("A3:F3").Merge
ws.Range("G3:H4").Merge
ws.Range("A7:H7").Merge
ws.Range("A4:C4").Merge
ws.Range("A5:C5").Merge
ws.Range("D4:F4").Merge
ws.Range("D5:F5").Merge
ws.Range("G5:H5").Merge

Dim i As Integer
i = 1
Dim row As Integer
'row keeps track of datasheet row
row = 2
Dim crow As Integer
'crow keeps track of credit report sheet row of table that holds upcoming transactions
crow = 9
Dim datedistance As Long
'datedistance used to check if the transaction is relevant (today or will happen in future)
datedistance = 0

'Looping through Purchases and adding relevant transactions to Credit report worksheet
Do While data.Cells(i, 2).Value <> ""
    If data.Range("B" & row) = "PURCHASES" Then
        If data.Range("F" & row) = acc Then
            datedistance = datediff("d", today, FormatDateTime(data.Range("U" & row), vbShortDate))
            If datedistance >= 0 Then
                ws.Range("A" & crow) = FormatDateTime(data.Range("A" & row), vbShortDate)
                ws.Range("B" & crow) = data.Range("H" & row)
                ws.Range("C" & crow) = data.Range("J" & row)
                ws.Range("D" & crow) = data.Range("O" & row)
                ws.Range("E" & crow) = data.Range("X" & row)
                ws.Range("F" & crow) = data.Range("AJ" & row)
                ucredit = ucredit + ws.Range("F" & crow)
                ws.Range("G" & crow) = icredit - ucredit
                ws.Range("H" & crow) = data.Range("U" & row)
                crow = crow + 1
                
            End If
        End If
    End If
    row = row + 1
    i = i + 1
Loop
'Adding commas and 3 decimal points
ws.Range("D9:F500").Select
Selection.NumberFormat = "#,##0.000"
Selection.FormatConditions.Add Type:=xlTextString, String:=".", _
    TextOperator:=xlContains
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
Selection.FormatConditions(1).StopIfTrue = True
'Adding commas and 3 decimal points
ws.Range("G5").Select
Selection.NumberFormat = "#,##0.000"
Selection.FormatConditions.Add Type:=xlTextString, String:=".", _
    TextOperator:=xlContains
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
Selection.FormatConditions(1).StopIfTrue = True
ws.Range("A2").Select
'Adding commas and 3 decimal points
ws.Range("G9:H500").Select
Selection.NumberFormat = "#,##0.000"
Selection.FormatConditions.Add Type:=xlTextString, String:=".", _
    TextOperator:=xlContains
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
Selection.FormatConditions(1).StopIfTrue = True
ws.Range("A2").Select
'Hiding blank/empty rows
ws.Range("A24:A500").AutoFilter 1, "<>", , , False
ws.Range("A5") = ucredit
ws.Range("D5") = icredit - ucredit
Application.ScreenUpdating = True
End Sub
