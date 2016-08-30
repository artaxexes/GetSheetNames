Option Explicit

Const sourceFolder = "C:\in\"
Const resultFolder = "C:\out\"

Dim sheets : sheets = GetSheetNames(ListXlsInFolder(sourceFolder), sourceFolder)
Dim i
For i = 0 To UBound(sheets)
  MsgBox sheets(i)(0) & ": " & vbCrLf & Join(sheets(i)(1), vbCrLf)
Next

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
