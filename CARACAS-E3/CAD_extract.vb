Option Explicit

' --- CONFIG ---
Private Const COL_TEXT As Long = 1 ' A
Private Const COL_X As Long = 2    ' B
Private Const COL_Y As Long = 3    ' C

' Tune these
Private Const X_TOL As Double = 50     ' same unit as your extraction
Private Const Y_TOL As Double = 50

Public Sub GroupTextByEquipment()
    Dim wsIn As Worksheet, wsOut As Worksheet
    Set wsIn = ThisWorkbook.Worksheets("Extract")
    Set wsOut = EnsureSheet("Output")

    Dim lastRow As Long
    lastRow = wsIn.Cells(wsIn.Rows.Count, COL_TEXT).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' Output headers
    wsOut.Cells.Clear
    wsOut.Range("A1:E1").Value = Array("EquipmentID", "AnchorX", "AnchorY", "NearbyText", "NearbyCount")

    Dim outRow As Long: outRow = 2
    Dim r As Long

    For r = 2 To lastRow
        Dim txt As String
        txt = CleanText(CStr(wsIn.Cells(r, COL_TEXT).Value))

        If IsEquipmentTag(txt) Then
            Dim x0 As Double, y0 As Double
            x0 = CDbl(wsIn.Cells(r, COL_X).Value)
            y0 = CDbl(wsIn.Cells(r, COL_Y).Value)

            Dim collected As Collection
            Set collected = New Collection

            Dim r2 As Long
            For r2 = 2 To lastRow
                If r2 <> r Then
                    Dim t2 As String
                    t2 = CleanText(CStr(wsIn.Cells(r2, COL_TEXT).Value))
                    If Len(t2) > 0 Then
                        Dim x As Double, y As Double
                        x = CDbl(wsIn.Cells(r2, COL_X).Value)
                        y = CDbl(wsIn.Cells(r2, COL_Y).Value)

                        If Abs(x - x0) <= X_TOL And Abs(y - y0) <= Y_TOL Then
                            SafeAddUnique collected, t2
                        End If
                    End If
                End If
            Next r2

            wsOut.Cells(outRow, 1).Value = txt
            wsOut.Cells(outRow, 2).Value = x0
            wsOut.Cells(outRow, 3).Value = y0
            wsOut.Cells(outRow, 4).Value = JoinCollection(collected, " | ")
            wsOut.Cells(outRow, 5).Value = collected.Count

            outRow = outRow + 1
        End If
    Next r

    wsOut.Columns("A:E").AutoFit
End Sub

' --- Helpers ---

Private Function IsEquipmentTag(ByVal s As String) As Boolean
    ' Simple patterns:
    ' T followed by digits, S followed by digits, etc.
    ' Add more patterns as needed.
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.IgnoreCase = True
    re.Global = False

    ' Examples:
    ' T12, S3, CB10, F1, etc.
    re.Pattern = "^(T|S|CB|F)\d+$"

    IsEquipmentTag = re.Test(s)
End Function

Private Function CleanText(ByVal s As String) As String
    s = Trim$(s)
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    CleanText = s
End Function

Private Sub SafeAddUnique(ByRef c As Collection, ByVal item As String)
    ' Prevent duplicates using a key
    On Error GoTo AlreadyThere
    c.Add item, LCase$(item)
    Exit Sub
AlreadyThere:
    On Error GoTo 0
End Sub

Private Function JoinCollection(ByVal c As Collection, ByVal delim As String) As String
    Dim i As Long, arr() As String
    If c.Count = 0 Then
        JoinCollection = ""
        Exit Function
    End If
    ReDim arr(1 To c.Count)
    For i = 1 To c.Count
        arr(i) = CStr(c(i))
    Next i
    JoinCollection = Join(arr, delim)
End Function

Private Function EnsureSheet(ByVal name As String) As Worksheet
    On Error Resume Next
    Set EnsureSheet = ThisWorkbook.Worksheets(name)
    On Error GoTo 0
    If EnsureSheet Is Nothing Then
        Set EnsureSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        EnsureSheet.Name = name
    End If
End Function
