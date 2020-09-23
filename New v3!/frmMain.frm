VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "File Dependency Sniffer v3  by Mike3dd"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3255
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmeLayout 
      Height          =   3280
      Left            =   10
      TabIndex        =   0
      Top             =   -50
      Width           =   6540
      Begin VB.Frame frmeOpt 
         Caption         =   "Needed Information:"
         Height          =   1590
         Left            =   150
         TabIndex        =   2
         Top             =   150
         Width           =   6240
         Begin VB.CheckBox chkSort 
            Caption         =   "Sort"
            Height          =   240
            Left            =   1875
            TabIndex        =   15
            Top             =   1050
            Width           =   1065
         End
         Begin VB.CheckBox chkKillDupes 
            Caption         =   "Kill Dupes"
            Height          =   240
            Left            =   1875
            TabIndex        =   14
            Top             =   1275
            Width           =   1065
         End
         Begin VB.CommandButton cmdBrowse 
            Caption         =   "Browse..."
            Default         =   -1  'True
            Height          =   315
            Left            =   5250
            TabIndex        =   11
            Top             =   1080
            Width           =   825
         End
         Begin VB.TextBox txtPath 
            Height          =   285
            Left            =   750
            OLEDropMode     =   1  'Manual
            TabIndex        =   10
            Top             =   720
            Width           =   5340
         End
         Begin VB.PictureBox Picture1 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   480
            Left            =   120
            Picture         =   "frmMain.frx":058A
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   9
            Top             =   360
            Width           =   480
         End
         Begin VB.CommandButton cmdExtract 
            Caption         =   "Extract"
            Height          =   315
            Left            =   4500
            TabIndex        =   8
            Top             =   1080
            Width           =   705
         End
         Begin VB.TextBox txtExt 
            Height          =   285
            Left            =   4800
            TabIndex        =   7
            Text            =   "*.dll, *.ocx, *.exe"
            Top             =   300
            Width           =   1290
         End
         Begin VB.CommandButton cmdClear 
            Caption         =   "Clear"
            Height          =   315
            Left            =   3750
            TabIndex        =   6
            Top             =   1080
            Width           =   705
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "Save"
            Height          =   315
            Left            =   3000
            TabIndex        =   5
            Top             =   1080
            Width           =   705
         End
         Begin VB.CheckBox chkFilter 
            Caption         =   "Use Filter"
            Height          =   240
            Left            =   750
            TabIndex        =   4
            Top             =   1275
            Width           =   1290
         End
         Begin VB.CheckBox chkAllFiles 
            Caption         =   "All Files"
            Height          =   195
            Left            =   750
            TabIndex        =   3
            Top             =   1050
            Width           =   1440
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "File to check:"
            Height          =   255
            Left            =   720
            TabIndex        =   13
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label lblExt 
            BackStyle       =   0  'Transparent
            Caption         =   "Extentions:"
            Height          =   255
            Left            =   3825
            TabIndex        =   12
            Top             =   375
            Width           =   915
         End
      End
      Begin VB.ListBox lstFile 
         Height          =   1230
         Left            =   150
         TabIndex        =   1
         Top             =   1905
         Width           =   6255
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Started on 02/14/02 -  Updated (v3) on 02/18/02

'--------------------------------
'  File Dependency Sniffer v3
'  Written by Mike Canejo
'--------------------------------

'AIM (aol): Mike3dd
'Email: MikeCanejo@hotmail.com

'Enjoy!


Option Explicit

Dim sOldOpenPath As String 'To remember dialog open from path
Dim sOldSavePath As String 'To remember dialog save to path



Private Sub chkAllFiles_Click()
    txtExt.Enabled = chkAllFiles.Value - 1
End Sub

Private Sub cmdExtract_Click()
Dim i As Integer
Dim X As Integer
Dim Z As Integer
Dim iFind As Long
Dim sExt As String
Dim iLen As Integer
Dim sFile As String
Dim iFree As Integer
Dim sFound As String
Dim sQuery As String
Dim bValid As Boolean
Dim iTerminator(1) As Long  'No need for more vars,
                            'so i used iTerminator
                            'as a array, my preference...


    'On Error GoTo ExitSearch
    

    
    iFree = FreeFile                                     'Get an unused file number
    Open txtPath For Binary Access Read As #iFree        'Opens the file for reading
        sFile = Space(LOF(iFree))                        'Puts the null terminator at the end of string variable
        Get #iFree, , sFile                              'Can now put the file into the variable cause it has space to accommodate it
    Close #iFree                                         'Close the process
    
    If Me.Height > 2330 Then
        Me.Tag = Me.Height
        Me.Height = 2330                                                'Hide list from manipulative tasks
    Else
        Me.Height = 4500
    End If
    
    sFile = LCase(sFile)                                                'To prevent search ambiguity, make it non case sensitive
    txtExt = LCase(txtExt)                                              'Same thing
    
    If chkAllFiles.Value Then
        'Search for all files in the Executable.
        Do
            DoEvents
            iFind = InStr(iFind + 1, sFile, ".")                        'Find the "." placement in the string
            If iFind = 0 Then Exit Do
            iTerminator(0) = InStrRev(sFile, Chr(0), iFind)             'Chr(0) is used to determinate the beginning of the file found and the ending
            'iTerminator(0) = pInstrRev(sFile, Chr(0), iFind)           'For VB5 Users, comment the one above and uncomment this line
            iTerminator(1) = iFind + 4                                  'To determine the ending, ifind is the start of say for example ".dll", +4 cause theres 4 letters so it gets the ending
            If iTerminator(0) And Mid$(sFile, _
            iTerminator(1), 1) = Chr$(0) Then                            'Beginning point and end point of file in string
                If iTerminator(1) - iTerminator(0) - 1 < 20 _
                And iTerminator(1) - iTerminator(0) - 1 > 5 Then        'Some parameters to make sure the findings are not something other than files
                                                                        'This assumes all dlls are less than a length of 20 and greater than 1 char length.
                    bValid = True
                    
                    sFound = Mid$(sFile, iTerminator(0) + 1, _
                    iTerminator(1) - iTerminator(0) - 1)
                    
                    If chkFilter.Value _
                    And InStr(sFound, Chr$(0)) = 0 _
                    And Len(Left(sFound, InStr(sFound, ".") - 1)) > 1 Then
                        If isFunky(Mid$(sFile, iTerminator(1) - 4, 4)) Then
                                bValid = False
                            Else
                                bValid = isFilename(sFound, Mid$(sFile, iTerminator(1) - 4, 4))
                        End If
                    Else
                        If chkFilter.Value Then bValid = False
                    End If
                    
                    If bValid Then
                        lstFile.AddItem sFound                          'If not detected in filters then add it
                    End If
                    
                    Debug.Print sFound                                  'Display in immediate window
                    
                    End If
            End If
        Loop
    Else
        'Searches for all files with the
        'extentions found in the txtExt textbox.
    
        For i = 1 To CharsIN(txtExt, "*")                                   'Search for each ext in query
            sExt = Mid$(txtExt, CharsPOS(txtExt, "*", i) + 1, 4)
            Do
                DoEvents
                iFind = InStr(iFind + 1, sFile, sExt)                       'Find the file extention in the string
                If iFind = 0 Then Exit Do
                iTerminator(0) = InStrRev(sFile, Chr(0), iFind)             'Chr(0) is used to determinate the beginning of the file found and the ending
                'iTerminator(0) = pInstrRev(sFile, Chr(0), iFind)           'For VB5 Users, comment the one above and uncomment this line
                iTerminator(1) = iFind + 4                                  'To determine the ending, ifind is the start of say for example ".dll", +4 cause theres 4 letters so it gets the ending
                If iTerminator(0) And Mid$(sFile, _
                iTerminator(1), 1) = Chr$(0) Then                            'Beginning point and end point of file in string
                    If iTerminator(1) - iTerminator(0) - 1 < 20 _
                    And iTerminator(1) - iTerminator(0) - 1 > 5 Then        'Some parameters to make sure the findings are not something other than files
                                                                            'This assumes all dlls are less than a length of 20 and greater than 1 char length.
                        bValid = True
                        
                        sFound = Mid$(sFile, iTerminator(0) + 1, _
                        iTerminator(1) - iTerminator(0) - 1)
                        
                        If chkFilter.Value Then
                            bValid = isFilename(sFound, sExt)
                        End If
                        
                        If bValid Then
                            lstFile.AddItem sFound                       'If not detected in filters then add it
                        End If
                        
                        Debug.Print sFound                                  'Display in immediate window
                    End If
                End If
            Loop
            iFind = 0
            
        Next i
    End If
    
    If chkKillDupes.Value Then KillDupesAPI lstFile                      'Remove doubles found from search
    If chkKillDupes.Value Then Sort lstFile, True                        'Sorts the listbox alphabetically
    
    If lstFile.ListCount > 0 Then lstFile.ListIndex = 0        'Set listindex to 0 in listbox
    If Me.Height < 2350 Then Me.Height = Me.Tag                'Resize form to show listbox
    
    Exit Sub
ExitSearch:
    MsgBox "Error#" & Err & " - " & Error(Err), vbCritical, "Error Occured"
End Sub

Private Function isFilename(sFilename As String, sExtention As String) As Boolean
On Error Resume Next
    '                       bValid = isFilename(sFound, sExt)
    '                       sFilename points to sFound and sExtention points to sExt
                        
    'NOTE: sFilename is a pointer to a var in reference to it from the function syntax.
    'So if the contents of any pointer changes, then the var its pointing
    'to does as well.. since its really sFound with a different name in memory.
    'The search code above uses this function to check
    'a found filename using var sFound so when sFilename is changed anywhere below,
    'sFound, which its pointing to, does as well... I'm just pointing this
    'out for newcommers because this is something that took me a while
    'to figure out when i just started to write functions in vb way back when
    'and if your learning vb on your own like i did then this is something you
    'figure out by trial and error usually... c++ helped me as well..but that's
    'a different story.
    
    'okie dokie
    '-Mike Canejo  ;]
     
    Dim i As Integer, X As Integer
    Dim iLen As Integer
    
    If InStr(sFilename, "\") Then
        sFilename = Mid(sFilename, _
        InStr(sFilename, "\") + 1)      'Sometimes "\" are in the filename
                                        'cause of paths in the exe.
                                        'So this will get the right of it.
    End If
    
    
    
                                        'Two loops below to detect a funky char in the filename.
                                        'If one is found it will cut the string off
                                        'at the pos its found cause almost every time
                                        'it's a wrong beginning of the filename being found.
                                        'The ending is always correct on the Chr(0) finding.
                                        'So this so far, as I can see, takes care of it...
                                        
                                        'Please leave feedback if I am wrong!
                                        
                                        'Again-
                                        'AIM (aol): Mike3dd
                                        'E-mail: MikeCanejo@hotmail.com
                    
                    
    For i = Len(sFilename) To 1 Step -1             'Start searching from end of string to beginning
        For X = 1 To 39
            If Mid$(sFilename, i, 1) = Chr$(X) _
            Or Mid$(sFilename, i, 1) = Chr$(96) Then
                sFilename = Mid$(sFilename, i + 1)   'Funky char found, cut it at the pos
                Exit For                            'Exit the loop since it found it
            End If
        Next X
    Next i
    For i = Len(sFilename) To 1 Step -1             'Start searching from end of string to beginning
        For X = 123 To 255
            If Mid$(sFilename, i, 1) = Chr$(X) _
            Or Mid$(sFilename, i, 1) = Chr$(96) Then
                sFilename = Mid$(sFilename, i + 1)   'Funky char found, cut it at the pos
                Exit For                            'Exit the loop since it found it
            End If
        Next X
    Next i
    
    iLen = Len(Left(sFilename, InStr( _
    sFilename, sExtention) - 1))        'Length parameters to ensure the filtered filename
                                        'is considered a "correct" file name length.
                                        'You can change this to your own liking...
                                                                        
    If iLen < 20 And iLen > 1 Then
        isFilename = True
    Else
        isFilename = False
    End If

End Function

Private Function isFunky(sCheck As String) As Boolean
'On Error Resume Next
    '                       bValid = isFilename(sFound, sExt)
    '                       sFilename points to sFound and sExtention points to sExt
                        
    'NOTE: sFilename is a pointer to a var in reference to it from the function syntax.
    'So if the contents of any pointer changes, then the var its pointing
    'to does as well.. since its really sFound with a different name in memory.
    'The search code above uses this function to check
    'a found filename using var sFound so when sFilename is changed anywhere below,
    'sFound, which its pointing to, does as well... I'm just pointing this
    'out for newcommers because this is something that took me a while
    'to figure out when i just started to write functions in vb way back when
    'and if your learning vb on your own like i did then this is something you
    'figure out by trial and error usually... c++ helped me as well..but that's
    'a different story.
    
    'okie dokie
    '-Mike Canejo  ;]
     
    Dim i As Integer, X As Integer
    
    
    
    
                                        'Two loops below to detect a funky char in the filename.
                                        'If one is found it will cut the string off
                                        'at the pos its found cause almost every time
                                        'it's a wrong beginning of the filename being found.
                                        'The ending is always correct on the Chr(0) finding.
                                        'So this so far, as I can see, takes care of it...
                                        
                                        'Please leave feedback if I am wrong!
                                        
                                        'Again-
                                        'AIM (aol): Mike3dd
                                        'E-mail: MikeCanejo@hotmail.com
                    

    For i = Len(sCheck) To 1 Step -1             'Start searching from end of string to beginning
        For X = 1 To 39
            If Mid(sCheck, i, 1) = Chr(X) Then
                isFunky = True  'Funky char found, cut it at the pos
                Exit Function                         'Exit the loop since it found it
            End If
        Next X
    Next i
    
    
    For i = Len(sCheck) To 1 Step -1             'Start searching from end of string to beginning
        For X = 123 To 255
            If Mid(sCheck, i, 1) = Chr(X) Then
                isFunky = True
                Exit Function                        'Exit the loop since it found it
            End If
        Next X
    Next i
    

End Function

Private Function CharsIN(sText As String, sChar As String) As Long
    'Wrote this to find the amount
    'of extentions to query in search.
    'Rather useful function too....
    
    Dim iPos As Long, sNext As String
    sNext = sText
    Do
        iPos = InStr(sText, sChar)
        If iPos = 0 Then Exit Function
        sText = Mid(sText, iPos + 1)
        CharsIN = CharsIN + 1
    Loop
End Function

Private Function CharsPOS(sText As String, sChar As String, Optional ByVal iStart As Long = 1) As Long
    'Wrote this to get a position of a char
    'found at a certain amount of times, get the pos
    
    Dim iPos As Long, iCount As Long
    iCount = 1
    Do

        iPos = InStr(iPos + 1, sText, sChar)
        If iPos = 0 Then Exit Function
        If iCount = iStart Then
            CharsPOS = iPos
            Exit Do
        End If
        iCount = iCount + 1
    Loop
End Function

Private Sub cmdBrowse_Click()
    On Error GoTo DialogError
    dFileName = vbNullString
    If sOldOpenPath = vbNullString Then sOldOpenPath = App.Path
    Dialog "Executable(*.exe)" + Chr$(0) + "*.exe" + Chr$(0) + "All Files " _
    & "(*.*)" + Chr$(0) + "*.*" + Chr$(0), "Search File", Me, ".exe", sOldOpenPath, True
    'Dialog is a function i wrote in the modDialog module, it's great!
    'i reuse this module in all my projects, to prevent usage of the
    'CommonDialog control of course, since this is alot better.
    
    If dFileName = vbNullString Then Exit Sub
    sOldOpenPath = dFileName
    txtPath = dFileName
    Exit Sub
DialogError:
    MsgBox "Error#" & Err & " - " & Error(Err), vbCritical, "Error Occured"
End Sub

Private Sub cmdSave_Click()
    On Error GoTo DialogError
    dFileName = vbNullString
    If sOldSavePath = vbNullString Then sOldSavePath = App.Path
    Dialog "Text File(*.txt)" + Chr$(0) + "*.txt" + Chr$(0) + "Document File " _
    & "(*.doc)" + Chr$(0) + "*.doc" + Chr$(0), "Save Dependency List", Me, ".txt", sOldSavePath, False
    If dFileName = vbNullString Then Exit Sub
    sOldSavePath = dFileName
    SaveListBox lstFile, dFileName, True
    Exit Sub
DialogError:
    MsgBox "Error#" & Err & " - " & Error(Err), vbCritical, "Error Occured"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 7000 Then Me.Width = 7000
    If Me.Height < 2331 Then Me.Height = 2330
    frmeLayout.Width = Me.Width - 140
    frmeLayout.Height = Me.Height - 440
    
    lstFile.Width = Me.Width - 430
    lstFile.Height = Me.Height - 2400

    frmeOpt.Width = lstFile.Width
    txtPath.Width = frmeOpt.Width - 920
    txtExt.Left = txtPath.Width - 540
    lblExt.Left = txtExt.Left - lblExt.Width - 50
    cmdBrowse.Left = txtPath.Width - 90
    cmdExtract.Left = cmdBrowse.Left - cmdExtract.Width - 100
    cmdClear.Left = cmdExtract.Left - cmdClear.Width - 100
    cmdSave.Left = cmdClear.Left - cmdSave.Width - 100
    'Some form resizing routines, nothing much
End Sub

Private Sub lstFile_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
    If Button = 2 Then
        Clipboard.Clear
        Clipboard.SetText lstFile.List(lstFile.ListIndex)
        'Copy the file name to clipboard from right clicking the listbox
    End If
End Sub

Private Sub cmdClear_Click()
    lstFile.Clear 'Remove all items in the listbox
End Sub

Private Sub txtPath_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    txtPath = Data.Files(1)
    'Drag and drop a file into the txtpath box instead of using browse feature, your choice
End Sub

Private Sub Form_Load()
    Me.Height = 2330
End Sub

