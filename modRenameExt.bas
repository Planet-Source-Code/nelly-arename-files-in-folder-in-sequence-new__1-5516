Attribute VB_Name = "modRenameExt"

Dim lngSCase As Long

Option Explicit

'renames extension only
Public Sub subReanmeExt()

    'On Error Resume Next

    Dim strCont(1) As String
    Dim intReCounter As Integer
    Dim pt As Integer
    Dim x As Integer
    Dim Y As Integer
    Dim strRename As String
    Dim strReExtension As String
    
    strCont$(0) = MsgBox("Are You Sure You Want To Continue", vbInformation + vbYesNo)
        If strCont$(0) = vbNo Then
        Exit Sub
        Else

       intReCounter% = frmGlobal.txtRnameCount.Text
       strReExtension$ = frmGlobal.txtReExt.Text

       'file count progressbar
       frmGlobal.pbrFiles.Max = 101
       pt% = frmGlobal.SBList1.ListCount - 1
    
       frmGlobal.File1.ListIndex = 0
       frmGlobal.SBList1.ListIndex = 0

       For x = 0 To frmGlobal.File1.ListCount - 1 'This is the code that renames the selected file
            On Error GoTo errName
        
           'set listindex
           frmGlobal.File1.ListIndex = (x)
           frmGlobal.SBList1.ListIndex = (x)

           'set file progressbar index
           Y = frmGlobal.SBList1.ListIndex / pt% * Val(100)

           If frmGlobal.SBList1.Selected(x) Then
               On Error GoTo errName
               strRename$ = ""
               strRename$ = Left(frmGlobal.File1.FileName, Len(frmGlobal.File1.FileName) - 4)
               
               'names selected files
               If frmGlobal.chkLCase.Value = vbChecked Then
                   Name modGlobalSearch.GetSelectedFile(frmGlobal.Dir1.Path) As modGlobalSearch.GetSelectedFileD(frmGlobal.Dir1.Path) & LCase(strRename$ & strReExtension$)
                ElseIf frmGlobal.chkUCase.Value = vbChecked Then
                   Name modGlobalSearch.GetSelectedFile(frmGlobal.Dir1.Path) As modGlobalSearch.GetSelectedFileD(frmGlobal.Dir1.Path) & UCase(strRename$ & strReExtension$)
                ElseIf frmGlobal.chkSCase.Value = vbChecked Then
                   Name modGlobalSearch.GetSelectedFile(frmGlobal.Dir1.Path) As modGlobalSearch.GetSelectedFileD(frmGlobal.Dir1.Path) & ftnSelCase & strReExtension$
                Else
                   Name modGlobalSearch.GetSelectedFile(frmGlobal.Dir1.Path) As modGlobalSearch.GetSelectedFileD(frmGlobal.Dir1.Path) & strRename$ & strReExtension$
                End If
           
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

        'reset progressbar and caption
        frmGlobal.pbrFiles.Value = (0)
        frmGlobal.lblPB.Caption = ""
    
    End If
    
Exit Sub

errName:
    MsgBox "FileName Already Exists", vbCritical + vbOKOnly
    'reset progressbar and caption
    frmGlobal.pbrFiles.Value = (0)
    frmGlobal.lblPB.Caption = ""


End Sub

'this is the function that sets the first and every character after " " to Ucase
Public Function ftnSelCase() As String

    Dim strRename As String

    'load the file into a textbox after stripping the last three characters
    frmGlobal.Text1.Text = Left(frmGlobal.File1.FileName, Len(frmGlobal.File1.FileName) - 4)
    'set the text to Lcase
    frmGlobal.Text1.Text = LCase(frmGlobal.Text1.Text)
    'start at the first character
    frmGlobal.Text1.SelStart = 0
    'set the seltext to 1 character
    frmGlobal.Text1.SelLength = 1
    'make the seltext to ucase
    frmGlobal.Text1.SelText = UCase(frmGlobal.Text1.SelText)
    
    Do 'search through the string looking for blank spaces
        lngSCase& = InStr(lngSCase& + 1, frmGlobal.Text1.Text, " ", vbTextCompare)
            If lngSCase > 0 Then 'if it finds a blank space, select the next character
                                 'and set it to Ucase
                frmGlobal.Text1.SelStart = lngSCase
                frmGlobal.Text1.SelLength = 1
                frmGlobal.Text1.SelText = UCase(frmGlobal.Text1.SelText)
            End If
    DoEvents
    Loop Until lngSCase <= 0 'loop until end of string
    ftnSelCase = frmGlobal.Text1.Text 'return the function with the manipulated text
    frmGlobal.Text1.Text = "" 'clear the textbox
    
End Function

