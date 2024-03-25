Attribute VB_Name = "Module1"
Sub Find_Files()
    
    Dim fldr As FileDialog
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    fldr.Show
    
    f = fldr.SelectedItems(1)
    
    f = f & "\"

    ibox = InputBox("File Must Contain (Note * wildcards can be used)", , "*.xls*")
    On Error GoTo ext
    
    sn = Split(CreateObject("wscript.shell").exec("cmd /c Dir """ & f & ibox & """ /s /a /b").StdOut.ReadAll, vbCrLf)
     
    Sheets(1).Cells(2, 1).Resize(UBound(sn) + 1) = Application.Transpose(sn)

ext:
    formulaLink
End Sub


Function formulaLink()
    Columns("D:D").Select
    Selection.EntireColumn.Hidden = False
    
    'formula link fisier
    Worksheets("Output").Cells(1, 2) = "Link fisier"
    Worksheets("Output").Cells(2, 2) = "=HYPERLINK(A2,""LINK"")"
    Range("B2:B" & Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row).FillDown
    
    
    'extragere link folder
    Sheets("Formula").Select
    Range("D2").Select
    Selection.Copy
    Sheets("Output").Select
    Range("D2").Select
    ActiveSheet.Paste
    
    Range("D2:D" & Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row).FillDown
    
    'formula link folder
    Worksheets("Output").Cells(1, 3) = "Link Folder"
    Worksheets("Output").Cells(2, 3) = "=HYPERLINK(D2,""LINK"")"
    Range("C2:C" & Range("A" & ActiveSheet.Rows.Count).End(xlUp).Row).FillDown
    
    Columns("D:D").Select
    Selection.EntireColumn.Hidden = True
    
    Range("A1").Select
    
End Function
