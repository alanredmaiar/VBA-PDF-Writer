Attribute VB_Name = "Módulo1"

Sub dalecana()

    Dim vaArray     As Variant
    Dim i           As Integer
    Dim oFile       As Object
    Dim oFSO        As Object
    Dim oFolder     As Object
    Dim oFiles      As Object

    Set oFSO = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFSO.GetFolder(ThisWorkbook.Path)
    
''Folders

    Set oFolders = oFolder.subfolders

    If oFolders.Count = 0 Then Exit Sub
    

    ReDim vaArray(1 To oFolders.Count)
    i = 1
    For Each oFolderx In oFolders
        
        Name oFolderx As UCase(oFolderx)
    Next

''Files

'    If oFiles.Count = 0 Then Exit Function
'
'    ReDim vaArray(1 To oFiles.Count)
'    i = 1
'    For Each oFile In oFiles
'        vaArray(i) = oFile.Name
'        i = i + 1
'    Next
'
'    listfiles = vaArray


End Sub
