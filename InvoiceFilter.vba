Sub getInvoiceData()
'
' getInvoiceData Macro
' Deletes OUTPUT values and pastes current DATA to OUTPUT

'
    Sheets("OUTPUT").Select
    Range("A2:J2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Sheets("DATA").Select
    Range("A2:I2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Sheets("OUTPUT").Select
    Range("A2").Select
    ActiveSheet.Paste
End Sub
Sub convertToNumbers()
'
' convertToNumbers Macro converts the A column to numbers as SAP exports it as text
' This step is necessary for determineDirectDebit VLOOKUP function to work!

'
    Sheets("OUTPUT").Select
    Range("A2").Select
    [A:A].Select
    With Selection
        .NumberFormat = "General"
        .Value = .Value
    End With
    End Sub
Sub determineDirectDebit()
'
' determineDirectDebit Macro
'

'
    Sheets("OUTPUT").Select
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=VLOOKUP(RC[-9],'Direct Debit Accounts'!C[-9]:C[-7],3,0)"
    Range("J2").Select
    Selection.AutoFill Destination:=Range("J2:J" & Range("E" & Rows.Count).End(xlUp).Row)
    Range(Selection, Selection.End(xlDown)).Select
End Sub
Sub deleteRows()
'
' deleteRows Macro
'

'
    Sheets("OUTPUT").Select
    Range("J1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$J$13682").AutoFilter Field:=10, Criteria1:="Yes"
    Selection.Offset(1, 0).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.EntireRow.Delete
    Range("J1").Select
    Selection.AutoFilter
End Sub
Sub fullDDRemoval()
'
' fullDDRemoval Macro
'

'
    Call getInvoiceData
    Call convertToNumbers
    Call determineDirectDebit
    Call deleteRows
End Sub
