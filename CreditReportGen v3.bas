Attribute VB_Name = "Module1"
Sub CreditReportGenerator()
Attribute CreditReportGenerator.VB_ProcData.VB_Invoke_Func = " \n14"
'HongSeok (Sam) Baik for Global Energy Trading Ltd
'7/30/2018
'Macro created to generate credit summary report from a given data excel file

Application.ScreenUpdating = False
Dim ws As Worksheet
Dim data As Worksheet
Set data = Worksheets("DATA")
Dim crdata As Worksheet
Dim acc As String
acc = InputBox("For what account do you wish to create a credit report for?", "Input Account Name", "GUNVOR")
acc = UCase(acc)
Dim icredit As Long
icredit = 0
Dim ucredit As Long
ucredit = 0
Set ws = Sheets.Add
Dim today As Variant
today = InputBox("Enter the date from which you would like to see upcoming transactions:(DD/M/YY)", "Enter Date", Format(Date, "dd/m/yyyy"))
today = Format(today, "dd/m/yyyy")

'Getting credit line information for specified account
On Error GoTo CreditInput
'if credit data worksheet DNE then ask for user input
Set crdata = Worksheets("CREDIT DATA")
Dim crdatarange As Range
Set crdatarange = crdata.Range("A1:AZ1").Find("CREDIT LINE")
Dim craccrange As Range
Set craccrange = crdata.Range("A1:AZ1").Find("ACCOUNT")
If crdatarange Is Nothing Then
    MsgBox ("Please ensure you named the Excel worksheet with credit line information, CREDIT DATA. Please also ensure you named the Column header of credit line values, CREDIT LINE")
    icredit = InputBox("What is the initial credit line of " & acc & " ?", "Input Initial Credit", 5000000)
    Debug.Print "Conditions not met"
ElseIf craccrange Is Nothing Then
    MsgBox ("Please ensure you named the Excel worksheet with credit line information, CREDIT DATA. Please also ensure you named the Column header of account name values, ACCOUNT")
    icredit = InputBox("What is the initial credit line of " & acc & " ?", "Input Initial Credit", 5000000)
    Debug.Print "Conditions not met"
End If
Dim i As Integer
i = 2
Do While crdata.Cells(i, 1).Value <> ""
    If crdata.Cells(i, craccrange.Column) = acc Then
        icredit = crdata.Cells(i, crdatarange.Column)
    End If
    i = i + 1
Loop

If crdata Is Nothing Then
CreditInput:
    icredit = InputBox("What is the initial credit line of " & acc & " ?", "Input Initial Credit", 5000000)
End If
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
.ColumnWidth = .ColumnWidth * 1.6
End With
With ws.Columns("C")
.ColumnWidth = .ColumnWidth * 3
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
.ColumnWidth = .ColumnWidth * 3
End With
With ws.Columns("I")
.ColumnWidth = .ColumnWidth * 1.6
End With
'creating credit table
ws.Range("A3") = acc & " Credit Summary:"
ws.Range("G3") = "Initial Credit Line:"
ws.Range("A4") = "Credit used:"
ws.Range("D4") = "Credit available:"
ws.Range("A7") = "Upcoming Transactions beginning from " & Format(today, "dd/mmm/yyyy")
ws.Range("A8") = "TRAN DATE:"
ws.Range("B8") = "'S or P/NO:"
ws.Range("C8") = "BARGE:"
ws.Range("D8") = "GRADE:"
ws.Range("E8") = "QTY:"
ws.Range("F8") = "PRICE:"
ws.Range("G8") = "AMT:"
ws.Range("H8") = "CREDIT AVAILABLE:"
ws.Range("I8") = "DUE DATE:"
ws.Range("G5") = icredit
'formatting
ws.Range("A3:I5").Borders.LineStyle = xlContinuous
ws.Range("A7:I500").Borders.LineStyle = xlContinuous
ws.Range("A3:G3").Font.Bold = True
ws.Range("A4:D4").Font.Bold = True
ws.Range("A7").Font.Size = 13
ws.Range("A8:I8").Font.Bold = True
ws.Range("H8:I8").Interior.ColorIndex = 8
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
ws.Range("A9:A500").HorizontalAlignment = xlLeft
ws.Range("I9:I500").HorizontalAlignment = xlRight
ws.Range("A3:F3").Merge
ws.Range("G3:I4").Merge
ws.Range("A7:I7").Merge
ws.Range("A4:C4").Merge
ws.Range("A5:C5").Merge
ws.Range("D4:F4").Merge
ws.Range("D5:F5").Merge
ws.Range("G5:I5").Merge


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
            datedistance = datediff("d", today, data.Range("U" & row))
            If datedistance >= 0 Then
                ws.Range("A" & crow) = Format(data.Range("I" & row), "dd/mmm/yyyy")
                ws.Range("B" & crow) = data.Range("D" & row)
                ws.Range("C" & crow) = data.Range("H" & row)
                ws.Range("D" & crow) = data.Range("J" & row)
                ws.Range("E" & crow) = data.Range("O" & row)
                ws.Range("F" & crow) = data.Range("X" & row)
                ws.Range("G" & crow) = data.Range("AJ" & row)
                ucredit = ucredit + ws.Range("G" & crow)
                ws.Range("H" & crow) = "=G5-SUM(G9:G" & crow & ")"
                ws.Range("I" & crow) = Format(data.Range("U" & row), "dd/mmm/yyyy")
                
                crow = crow + 1
            End If
        End If
    End If
    row = row + 1
    i = i + 1
Loop
'Adding commas and 3 decimal points
ws.Range("D9:H500").Select
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
'Adding commas and 3 decimal points
ws.Range("A5:D5").Select
Selection.NumberFormat = "#,##0.000"
Selection.FormatConditions.Add Type:=xlTextString, String:=".", _
    TextOperator:=xlContains
Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
Selection.FormatConditions(1).StopIfTrue = True
ws.Range("A2").Select
'Hiding blank/empty rows
ws.Range("A20:A500").AutoFilter 1, "<>", , , False
ws.Range("A5") = "=SUM(G9:G500)"
ws.Range("D5") = "=G5-A5"

Dim more As Byte
'change more to 1 if you wish to add more columns to the Credit Report table, 0 if not
more = 0
If more = 1 Then
    i = 1
    row = 2
    crow = 9
    'add Column insertion code here below
    'example code: ws.Range("B:B").EntireColumn.Insert
    'example code: ws.Range("B8") = "TYPE REMKS"
    ws.Range("B:B").EntireColumn.Insert
    ws.Range("B8") = "TYPE REMKS"
    'new Loop to account for new columns if added
    Do While data.Cells(i, 6).Value <> ""
        If data.Range("B" & row) = "PURCHASES" Then
            If data.Range("F" & row) = acc Then
                datedistance = datediff("d", today, data.Range("U" & row))
                If datedistance >= 0 Then
                    'add data insertion code here below
                    'example code: ws.Range("B" & crow) = data.Range("C" & row)
                    ws.Range("B" & crow) = data.Range("C" & row)
                    
                    crow = crow + 1
                End If
            End If
        End If
        row = row + 1
        i = i + 1
    Loop
End If

Application.ScreenUpdating = True
End Sub
