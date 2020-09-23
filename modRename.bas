Attribute VB_Name = "modRename"
Dim sCase As String

Option Explicit

'This code renames the files in the selected directory
Public Sub subRename()

'On Error GoTo errRename

    Dim strCont(1) As String
    Dim intReCounter As Integer
    Dim strRename As String
    Dim strReExtension As String
    Dim x As Integer
    Dim pt As Integer
    Dim Y As Integer


    strCont$(0) = MsgBox("Are You Sure You Want To Continue", vbInformation + vbYesNo)
        If strCont$(0) = vbNo Then
        Exit Sub
        Else

            intReCounter% = frmGlobal.txtRnameCount.Text
            strRename$ = frmGlobal.txtRenameName.Text
            strReExtension$ = frmGlobal.txtReExt.Text

            frmGlobal.File1.ListIndex = 0
            frmGlobal.SBList1.ListIndex = 0
    
            'file count progressbar
            frmGlobal.pbrFiles.Max = 101
            pt% = frmGlobal.SBList1.ListCount - 1

                For x = 0 To frmGlobal.File1.ListCount - 1 'This is the code that renames the selected file
                    On Error GoTo errName
                    'set listindex
                    frmGlobal.File1.ListIndex = (x)
                    frmGlobal.SBList1.ListIndex = (x)
        
                    'set file progressbar index
                    Y = frmGlobal.SBList1.ListIndex / pt% * Val(100)
            
                    If frmGlobal.SBList1.Selected(x) Then
                       ' On Error GoTo errName
                        'names selected files
                        Name modGlobalSearch.GetSelectedFile(frmGlobal.Dir1.Path) As modGlobalSearch.GetSelectedFileD(frmGlobal.Dir1.Path) & strRename$ & intReCounter% & strReExtension$
                        'increase or decrease counter
                        If frmGlobal.chkIncrease.Value = vbChecked Then
                            intReCounter% = intReCounter% + 1
                        ElseIf frmGlobal.chkDecrease.Value = vbChecked Then
                            intReCounter% = intReCounter% - 1
                        End If
        
                    End If
        
                    'sets value of progress
                    frmGlobal.pbrFiles.Value = Val(Y)
                    frmGlobal.lblPB.Caption = Str(Y) & " %"
         

            DoEvents
            Next x
    
        'disply counters
        frmGlobal.txtReFileCount.Text = x
        frmGlobal.txtNamed.Text = Val(intReCounter%)
    
        frmGlobal.Dir1.Refresh 'Refresh directory when completed
        frmGlobal.subListFiles  'Refresh files in directory when completed
    
        strCont$(1) = MsgBox("Rename Completed", vbInformation + vbOKOnly) 'Tells you when all files have been Renamed

        'reste progressbar and caption
        frmGlobal.pbrFiles.Value = (0)
        frmGlobal.lblPB.Caption = ""
    
    End If

Exit Sub

errRename:
    MsgBox "Rename Unsuccessful", vbCritical + vbOKOnly
    'reset progressbar and caption
    frmGlobal.pbrFiles.Value = (0)
    frmGlobal.lblPB.Caption = ""
    

Exit Sub

errName:
    MsgBox "FileName Already Exists", vbCritical + vbOKOnly
    'reset progressbar and caption
    frmGlobal.pbrFiles.Value = (0)
    frmGlobal.lblPB.Caption = ""

End Sub


