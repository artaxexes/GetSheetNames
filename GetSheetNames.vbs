Option Explicit

Const sourceFolder = "C:\in\"
Const resultFolder = "C:\out\"

MsgBox Join(ListXlsInFolder(sourceFolder), vbCrLf)

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
