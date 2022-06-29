Sub NEW_TRANS()

Dim RWS As Integer
Dim dy As Integer
Dim st_dte As Date, end_dte As Date, WK_END As Date
Dim Q As Single, E As Single
Dim ite As String
Dim HOL As ListObject
Dim ASSGN As Single



Dim HOLS As ListObject

Set HOLS = Sheets("HOLIDAYS").ListObjects("HOLS")

RWS = Application.WorksheetFunction.CountA(Range("Gantt!A1:A5000"))

'MsgBox RWS
    With Worksheets("TRANS").ListObjects("TRANS_1")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Rows.Delete
        End If
    End With

For G = 2 To RWS






ite = Range("GANTT!A" & G).Value
st_dte = Range("GANTT!B" & G).Value
end_dte = Range("GANTT!C" & G).Value
ASSGN = Range("GANTT!D" & G).Value
'Now we need to loop through teh start and end dates.
ST_DY = Application.WorksheetFunction.Weekday(st_dte, 2)
END_DY = Application.WorksheetFunction.Weekday(end_dte, 2)

WKS = ((end_dte + 6 - END_DY) - (st_dte - ST_DY)) / 7
'HOL = HOLS.ListColumns(ite).DataBodyRange
'MsgBox WKS

 If HeaderExists("HOLIDAYS", "HOLS", ite) = True Then
    'MsgBox "You already added the Store calculation column " _
    '  & "to the table!"
    'Exit Sub
  Else
    'Take Some Action...
    HOLS.ListColumns.Add.Name = ite
    
    Range("HOLS[[HOL]]").Copy _
    Destination:=Range("HOLS[[" & ite & "]]")

    
  End If

'tbl.ListColumns.Add.Name = ITE1


For R = 1 To WKS

'TABLE = TRANS_1

If R = 1 Then
dy = Application.WorksheetFunction.Weekday(st_dte, 2)
WK_END = st_dte + 7 - dy
WK_STRT = st_dte - dy + 1

    If end_dte > WK_END Then
    Q = Application.WorksheetFunction.NetworkDays(st_dte, WK_END, HOLS.ListColumns(ite).DataBodyRange) * 8 * ASSGN
    E = Application.WorksheetFunction.NetworkDays(st_dte, WK_END, HOLS.ListColumns(ite).DataBodyRange) * 8
    Else
    Q = Application.WorksheetFunction.NetworkDays(st_dte, end_dte, HOLS.ListColumns(ite).DataBodyRange) * 8 * ASSGN
    E = Application.WorksheetFunction.NetworkDays(st_dte, end_dte, HOLS.ListColumns(ite).DataBodyRange) * 8
    End If

't = Application.WorksheetFunction.NetworkDays(st_dte, end_dte, Range(ite))
'MsgBox t

'IR = Insert_ROW(ite, WK_STRT, Q)

'ActiveSheet.ListObjects("Table1").ListRows.Add AlwaysInsert:= True

Else  ' IF R IS GREATER THAN 1

WK_STRT = WK_END + 1
'dy = Application.WorksheetFunction.Weekday(st_dte, 2)
WK_END = WK_STRT + 6

If WK_END < end_dte Then
Q = Application.WorksheetFunction.NetworkDays(WK_STRT, WK_END, HOLS.ListColumns(ite).DataBodyRange) * 8 * ASSGN
E = Application.WorksheetFunction.NetworkDays(WK_STRT, WK_END, HOLS.ListColumns(ite).DataBodyRange) * 8
Else
Q = Application.WorksheetFunction.NetworkDays(WK_STRT, end_dte, HOLS.ListColumns(ite).DataBodyRange) * 8 * ASSGN
E = Application.WorksheetFunction.NetworkDays(WK_STRT, end_dte, HOLS.ListColumns(ite).DataBodyRange) * 8

End If

End If

'dy = Application.WorksheetFunction.Weekday(st_dte, 2)

'WK_END = st_dte + 7 - dy
'Q = Application.WorksheetFunction.NetworkDays(st_dte, WK_END, Range(ite))


'MsgBox WK_END
IR = Insert_ROW(ite, WK_STRT, Q, E, st_dte, end_dte, WK_END)


Next R
'MsgBox dy

'Q = Application.WorksheetFunction.NetworkDays(st_dte, WK_END, Range(ite))



'MsgBox "DAYS = " & Q





Next G



End Sub


Function Insert_ROW(I_1, D_1, Q_1, E_1, SD, ED, WE)
Dim TableName As ListObject
Set TableName = Sheets("TRANS").ListObjects("TRANS_1")
Dim addedRow As ListRow
Set addedRow = TableName.ListRows.Add()
With addedRow
    .Range(1) = I_1
    .Range(2) = D_1
    .Range(3) = Q_1
    .Range(4) = E_1
    .Range(5) = SD
    .Range(6) = ED
    .Range(7) = WE
End With
End Function

Public Function HeaderExists(SHEETNAME As String, TableName As String, HeaderName As String) As Boolean

'PURPOSE: Output a true value if column name exists in specified table
'SOURCE: www.TheSpreadsheetGuru.com

Dim tbl As ListObject
Dim hdr As ListColumn

On Error GoTo DoesNotExist
  Set tbl = Sheets(SHEETNAME).ListObjects(TableName)
  Set hdr = tbl.ListColumns(HeaderName)
On Error GoTo 0

HeaderExists = True

Exit Function

'Error Handler
DoesNotExist:
  Err.Clear
  HeaderExists = False

End Function
