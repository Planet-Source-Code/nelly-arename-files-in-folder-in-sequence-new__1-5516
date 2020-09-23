Attribute VB_Name = "modGlobalSearch"
'Returns the name of the selected file in file1
Public Function GetSelectedFile(strPath As String) As String

If Right(strPath, 1) <> "\" Then
    GetSelectedFile = strPath & "\" & frmGlobal.File1.FileName
Else
    GetSelectedFile = strPath & frmGlobal.File1.FileName
    
End If
    
End Function

'Returns path of path only
Public Function GetSelectedFileD(strPath As String) As String

    If Right(strPath$, 1) <> "\" Then
        GetSelectedFileD = strPath$ & "\"
    Else
        GetSelectedFileD = strPath$
    End If

End Function
