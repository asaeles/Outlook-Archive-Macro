Attribute VB_Name = "modArchive"
Function GetArchiveFolder(MyMailItem As Outlook.MailItem) As String

    MyPath = MyMailItem.Parent.FolderPath
    Pos1 = InStr(3, MyPath, "\")
    GetArchiveFolder = "\\" & Format(MyMailItem.CreationTime, "yyyy") & Mid(MyPath, Pos1)

End Function

Function GetArchiveFolderItem(MyMailItem As Outlook.MailItem, Optional CreateIfNotFound As Boolean = True) As Outlook.Folder

    Dim MyCurrentFolder As Outlook.Folder
    Dim Check As Store
    Dim DefPath As String
    
    MyYear = Format(MyMailItem.CreationTime, "yyyy")
    'MyYear = "2013"
    MyPath = MyMailItem.Parent.FolderPath
    If InStr(MyPath, "\\" & MyYear) = 1 Then
        Err.Raise vbObjectError + 604, , "Already moved"
        Set GetArchiveFolderItem = Nothing
        Exit Function
    End If
    
    Pos1 = InStr(3, MyPath, "\")
    If Pos1 = 0 Then
        Err.Raise vbObjectError + 600, , "Source folder cannot be matched"
        Exit Function
    End If
    
    On Error Resume Next
    Set Check = Application.Session.Stores(MyYear)
    If Err.Number <> 0 Then
        Err.Clear
        On Error Resume Next
        Set Check = Application.Session.Stores(1)
        If Err.Number = 0 Then
            DefPath = Check.FilePath
        Else
            DefPath = "D:\"
        End If
        Err.Clear
        On Error GoTo 0
        MyRes = MsgBox("A new Outlook archive file need to be created to save the selected mail(s), create a new one?", vbOKCancel)
        If MyRes = vbOK Then
            MyFolder = GetFolder("Chose Outlook archive folder", DefPath)
            Application.Session.AddStore MyFolder & "\" & MyYear & ".pst"
            Application.GetNamespace("MAPI").Folders.GetLast.Name = MyYear
        Else
            Set GetArchiveFolderItem = Nothing
            Exit Function
        End If
    End If
    
    Set MyCurrentFolder = Application.GetNamespace("MAPI").Folders(MyYear)
    
    Do
        Pos2 = InStr(Pos1 + 1, MyPath, "\")
        If Pos2 = 0 Then
            Pos2 = Len(MyPath) + 1
        End If
        CurrentFolder = Mid(MyPath, Pos1 + 1, Pos2 - Pos1 - 1)
        On Error Resume Next
10      Set MyCurrentFolder = MyCurrentFolder.Folders(CurrentFolder)
        If Err.Number <> 0 Then
            Err.Clear
            On Error GoTo 0
            If CreateIfNotFound Then
                MyCurrentFolder.Folders.Add CurrentFolder
                GoTo 10
            Else
                Err.Raise vbObjectError + 603, , "Destination folder not found"
                Exit Function
            End If
        End If
        Pos1 = Pos2
    Loop While Pos2 < Len(MyPath)
    
    Set GetArchiveFolderItem = MyCurrentFolder

End Function

Public Sub ArchiveItems()

    Dim olMal As Outlook.MailItem
    Dim olSel As Collection
    Static MyErrCount, MyAlreadyMovedCount As Integer
    MyErrCount = 0
    MyAlreadyMovedCount = 0
    
    Set olSel = New Collection
    
    For i = 1 To Application.ActiveExplorer.Selection.Count
        olSel.Add Application.ActiveExplorer.Selection(i)
    Next i
    
    For i = 1 To olSel.Count
        On Error Resume Next
        Set olMal = olSel(i)
        If Err.Number = 0 Then
            If olMal.FlagRequest = "" And olMal.UnRead = False Then
                On Error Resume Next
                DoEvents
                olMal.Move GetArchiveFolderItem(olMal)
                DoEvents
                If Err.Number <> 0 Then
                    If Err.Number = vbObjectError + 604 Then
                        MyAlreadyMovedCount = MyAlreadyMovedCount + 1
                    Else
                        MyErrCount = MyErrCount + 1
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
        Else
            MyErrCount = MyErrCount + 1
            Err.Clear
            On Error GoTo 0
        End If
    Next i

    MyMsg = ""
    If MyErrCount > 0 Then MyMsg = MyErrCount & " error(s) encountered during move" & vbNewLine
    If MyAlreadyMovedCount > 0 Then MyMsg = MyMsg & MyAlreadyMovedCount & " item(s) already in place" & vbNewLine
    If MyMsg <> "" Then MsgBox MyMsg

End Sub

Private Sub Test()

   MsgBox GetArchiveFolderItem(Application.ActiveExplorer.Selection(1)).FolderPath

End Sub

