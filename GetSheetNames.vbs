Option Explicit

Const sourceFolder = "C:\in\"
Const resultFolder = "C:\out\"

SaveSheetNames GetSheetNames(ListXlsInFolder(sourceFolder), sourceFolder), resultFolder


' ListXlsInFolder
' receives: pathname of a folder
' returns: array with xls/xlsx file names
Private Function ListXlsInFolder(pathName)

  Dim files : files = Array()
  If FolderExists(pathName) Then
	  Dim objFSO : Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim objFolder : Set objFolder = objFSO.GetFolder(pathName)
    Dim objFiles : Set objFiles = objFolder.Files
	  Dim objFile
    For Each objFile In objFiles
      Dim extension : extension = objFSO.GetExtensionName(objFile)
	    If extension = "xls" Or extension = "xlsx" Then
        Dim index : index = UBound(files)
        ReDim Preserve files(index + 1)
        files(index + 1) = objFile.Name
		  End If
	  Next
  Else
	  MsgBox "This folder does not exist"
	End If

  ListXlsInFolder = files

End Function

' FolderExists
' receives: pathname of a folder
' returns: boolean
Private Function FolderExists(ByVal folderPath)

  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  FolderExists = fso.FolderExists(folderPath)
  Set fso = Nothing

End Function


' GetSheetNames
' receives: array with sheets names and pathname to folder
' returns: ragged multidimensional array (array of arrays)
Private Function GetSheetNames(sheetNames, path)

  Dim app : Set app = CreateObject("Excel.Application")
  app.DisplayAlerts = False

  Dim allNames : allNames = Array()
  Dim i, j
  For i = 1 To UBound(sheetNames)

    Dim wb : Set wb = app.Workbooks.Open(path & sheetNames(i), 0, True)
    Dim k, l
    Dim sheets : sheets = Array()
    For k = 1 To wb.Sheets.Count
      l = Ubound(sheets)
      ReDim Preserve sheets(l + 1)
      sheets(l + 1) = wb.Sheets(k).Name
    Next

    j = Ubound(allNames)
    ReDim Preserve allNames(j + 1)
    allNames(j + 1) = Array(sheetNames(i), sheets)

    wb.Saved = True
    wb.Close
    Set wb = Nothing

  Next

  app.Quit
  Set app = Nothing

  getSheetNames = allNames

End Function


' SaveSheetNames
' receives: ragged multidimensional array (array of arrays)
Private Sub SaveSheetNames(allNames, path)

    Dim app : Set app = CreateObject("Excel.Application")
    app.DisplayAlerts = False

    Dim wb : Set wb = app.Workbooks.Add
    Dim ws : Set ws = wb.Worksheets(1)
    app.Sheets(1).Select

    Dim i, j, k
    k = 1
    For i = 0 To UBound(allNames)
      For j = 0 To UBound(allNames(i)(1))
        ws.Cells(k, 1).Value = allNames(i)(0)
        ws.Cells(k, 2).Value = allNames(i)(1)(j)
        'MsgBox allNames(i)(0) & ": " & allNames(i)(1)(j)
        k = k + 1
      Next
    Next

    path = path & "SheetNames.xlsx"
    If FileExists(path) Then
      FileDelete(path)
    End If

    wb.SaveAs(path)
    wb.Close
    Set wb = Nothing
    Set ws = Nothing

End Sub


' FileExists
' receives: pathname
' returns: boolean
Private Function FileExists(ByVal filePath)

  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  FileExists = fso.FileExists(filePath)
  Set fso = Nothing

End Function

' FileDelete
' receives: file path
Private Sub FileDelete(filePath)

  Dim fso : Set fso = CreateObject("Scripting.FileSystemObject")
  fso.DeleteFile(filePath)
  Set fso = Nothing

End Sub
