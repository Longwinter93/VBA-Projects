Attribute VB_Name = "Module2"
Public Sub Clear()
    
    'clear all data from Worksheet
    ThisWorkbook.Worksheets(3).Range("N8:XFD1048576").Clear

End Sub

Sub UnzipAndUploadcsvfile()
    Dim FSO As Object
    Dim oApp As Object
    Dim Fname As Variant
    Dim FileNameFolder As Variant
    Dim DefPath As String
    Dim ruta As String
    ruta = ThisWorkbook.Path & "\"
    ChDir ruta
    

    Fname = Application.GetOpenFilename(FileFilter:="Zip Files (*.zip), *.zip", _
                                        MultiSelect:=False)
    If Fname = False Then
        Exit Sub
    Else
        'Destination folder
        DefPath = Application.ActiveWorkbook.Path    ' The original path of your  File
        If Right(DefPath, 1) <> "\" Then
            DefPath = DefPath & "\"
        End If

        FileNameFolder = DefPath

        '
        '
        '
        '

        'Extract the file into the file where your  File is
        Set oApp = CreateObject("Shell.Application")
        oApp.Namespace(FileNameFolder).CopyHere oApp.Namespace(Fname).items


        On Error Resume Next
        Set FSO = CreateObject("scripting.filesystemobject")
        FSO.deletefolder Environ("Temp") & "\Temporary Directory*", True
    End If



    Dim FileToOpen As Variant
    Dim OpenBook As Workbook


    Application.ScreenUpdating = False

    FileToOpen = Application.GetOpenFilename(Title:="Browse for your CSV File", FileFilter:="Excel Files (*.csv*),*csv*")
    If FileToOpen <> False Then
        Set OpenBook = Application.Workbooks.Open(FileToOpen)
        OpenBook.Sheets(1).Range("A1").CurrentRegion.Copy
        ThisWorkbook.Worksheets(3).Range("N8").PasteSpecial xlPasteValues
        
        Application.CutCopyMode = False
        
        
        OpenBook.Close False
        
         MsgBox "Task Done!"
    End If
    
    Application.ScreenUpdating = True


End Sub



