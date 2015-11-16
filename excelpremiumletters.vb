Sub PremiumLetters()

Dim ws As Worksheet
Dim ntable As Range
Dim firstcol As Range
Dim fname As Range
Dim lname As Range
'Dim position as range
Dim unit As Range
Dim Edu_pts_rng As Range
Dim lead_pts_rng As Range
Dim edu_pts_val As Integer
Dim lead_pts_val As Integer
Dim wd_template As Range

Dim rnum As Integer
Dim i As Integer

Dim objWord As Object

Set ws = ThisWorkbook.Worksheets("ForLetters")
Set ntable = ws.Range("Table1")
Set firstcol = ntable.Columns(1)

Dim qual_status As String
Dim not_qual As String
Dim prac_qual As String
Dim lead_qual As String
Dim both_qual As String

rnum = firstcol.Rows.Count

On Error Resume Next
Set objWord = GetObject(, "Word.Application")

    If objWord Is Nothing Then
        Set objWord = CreateObject("Word.Application")
    End If

    'It's good practice to reset error warnings
    On Error GoTo 0


For i = 1 To rnum

    edu_pts_val = 0
    lead_pts_val = 0
    
    Set fname = firstcol.Cells(i)
    Set lname = firstcol.Cells(i).Offset(0, 1)
'   position = firstcol.cells(i).offset(0,3).value
    Set unit = firstcol.Cells(i).Offset(0, 4)
    Set Edu_pts_rng = firstcol.Cells(i).Offset(0, 7)
    Set lead_pts_rng = firstcol.Cells(i).Offset(0, 8)
    Set wd_template = firstcol.Cells(i).Offset(0, 10)
    
    If Edu_pts_rng.Value = "N/A" Then
        edu_pts_val = 0
    ElseIf IsNumeric(Edu_pts_rng.Value) = True Then
       edu_pts_val = Edu_pts_rng.Value
    End If
    
    If lead_pts_rng.Value = "N/A" Then
        lead_pts_val = 0
    ElseIf IsNumeric(lead_pts_rng.Value) = True Then
        lead_pts_val = lead_pts_rng.Value
    End If
    
    If edu_pts_val < 70 And lead_pts_val < 60 Then
        wd_template.Value = "neither"
        qual_status = "not_qual"
    End If
    If edu_pts_val >= 70 And lead_pts_val < 60 Then
        wd_template.Value = "practice_only"
        qual_status = "prac_qual"
    End If
    If edu_pts_val < 70 And lead_pts_val >= 60 Then
        wd_template.Value = "lead_only"
        qual_status = "lead_qual"
    End If
    If edu_pts_val >= 70 And lead_pts_val >= 60 Then
        wd_template.Value = "both"
        qual_status = "both_qual"
    End If

'Debug.Print (fname.Value)
'Debug.Print (lname.Value)
'Debug.Print (unit.Value)
'Debug.Print (qual_status)

    objWord.documents.Open "C:\Users\smithem5\Desktop\Premium Letters\PremiumLetter2015Testing.dotm"
    objWord.Visible = True
    objWord.Run "CreatePremiumLetters", fname.Value, lname.Value, unit.Value, qual_status

    'Application.Wait (Now + TimeValue("0:00:01"))

Next i

End Sub
