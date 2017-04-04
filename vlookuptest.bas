Sub Finditem()

On Error Resume Next
Dim SFile As Variant
FPath = "C:\Users\me\Documents\VBAtestfile"
HPath = "C:\Users\me\Documents\VBA"

Dim OShell: Set OShell = CreateObject("Shell.Application")
Dim ODir: Set ODir = OShell.Namespace(FPath)

For Each SFile In ODir.items
 If InStr(ODir.Getdetailsof(SFile, 0), "Old") Then
 Else: Filename = ODir.Getdetailsof(SFile, 0)
 End If
Next

Hname = "vlookuptest.xlsx"
Workbooks.Open (HPath & "\" & Hname)
MsgBox ActiveSheet.Name
ActiveSheet.Range("D1").EntireColumn.Insert
Workbooks.Open (FPath & "\" & Filename)
Table1 = ActiveSheet.Range("A1:E99")
Workbooks(Hname).Activate
Table2 = ActiveSheet.Range("C1:C31")

Dim itm As Variant
i = 1
For Each itm In Table2
 Cells(i, 4).Value = Application.WorksheetFunction.VLookup(itm, Table1, 5, False)
 i = i + 1
Next

Workbooks(Filename).Close
Name FPath & "\" & Filename As FPath & "\" & Left(Filename, InStr(Filename, ".") - 1) & ".xlsx"

Workbooks(Hname).Save
Workbooks(Hname).Close

'MsgBox Application.WorksheetFunction.VLookup("A", Table2, 2, False)

End Sub