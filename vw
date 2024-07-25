#step 1.this macro is a revision of Jun's original code enabling automatic deletion of sheets that have 'na" while keeping the others 

Sub MacroNA()
Dim ws_count As Integer
Dim I As Integer
ws_count = ActiveWorkbook.Worksheets.Count
ws_start = ActiveWorkbook.ActiveSheet.Index
Application.DisplayAlerts = False
For I = ws_start To ws_count
   ActiveWorkbook.Worksheets(I).Activate
   ActiveWorkbook.Worksheets(I).Calculate 
   If IsError(ActiveWorkbook.Worksheets(I).Cells(6, 5)) Then
            ActiveWorkbook.Worksheets(I).Delete
            I = I - 1
   Else
            Cells.Select
            Selection.Copy
            Cells.Select
            Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End If        
Next I    
End Sub
··························································································································································································································
#step2 this macro can be used to delete the CN** sheets before the "index" sheet
Sub Macrodelete()
Dim I As Integer
J = 100 
I = 1
Application.DisplayAlerts = False
While J > 1
   ActiveWorkbook.Worksheets(I).Delete
   J = ActiveWorkbook.Worksheets("Index").Index
Wend
End Sub
.......................................................................................................................
#step3 this macro is used to write the filename of each sheet into a single sheet (the "index" sheet)'s Column D
Sub Macroname()
Dim ws As Worksheet
Dim x As Integer
x = 1
For Each ws In Activeworkbook.Worksheets
     Sheets("Index").Cells(x, 4) = right(ws.Name,4)
     x = x + 1
Next ws
End Sub
