VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "File Dependency Sniffer by Mike3dd"
   ClientHeight    =   3165
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3165
   ScaleWidth      =   5820
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFile 
      Height          =   1230
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   5580
   End
   Begin VB.Frame frmOpt 
      Caption         =   "Needed Information:"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5565
      Begin VB.CheckBox chkDupes 
         Caption         =   "Remove Dupes"
         Height          =   195
         Left            =   750
         TabIndex        =   12
         Top             =   1280
         Width           =   1440
      End
      Begin VB.CheckBox chkAppend 
         Caption         =   "Append to file"
         Height          =   240
         Left            =   750
         TabIndex        =   11
         Top             =   1050
         Width           =   1290
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   315
         Left            =   2325
         TabIndex        =   10
         Top             =   1080
         Width           =   705
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   315
         Left            =   3075
         TabIndex        =   9
         Top             =   1080
         Width           =   705
      End
      Begin VB.TextBox txtExt 
         Height          =   285
         Left            =   4800
         TabIndex        =   8
         Text            =   ".dll"
         Top             =   300
         Width           =   615
      End
      Begin VB.CommandButton cmdExtract 
         Caption         =   "Extract"
         Height          =   315
         Left            =   3825
         TabIndex        =   6
         Top             =   1080
         Width           =   705
      End
      Begin VB.PictureBox Picture1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   120
         Picture         =   "frmMain.frx":058A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   4
         Top             =   360
         Width           =   480
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   750
         OLEDropMode     =   1  'Manual
         TabIndex        =   2
         Top             =   720
         Width           =   4665
      End
      Begin VB.CommandButton cmdBrowse 
         Caption         =   "Browse..."
         Default         =   -1  'True
         Height          =   315
         Left            =   4575
         TabIndex        =   1
         Top             =   1080
         Width           =   825
      End
      Begin VB.Label lblExt 
         BackStyle       =   0  'Transparent
         Caption         =   "Extention:"
         Height          =   255
         Left            =   4050
         TabIndex        =   7
         Top             =   375
         Width           =   690
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "File to check:"
         Height          =   255
         Left            =   720
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by Mike Canejo
'AIM: Mike3dd
'Email: MikeCanejo@hotmail.com

'Enjoy!


Option Explicit

Dim sOldOpenPath As String 'To remember dialog open from path
Dim sOldSavePath As String 'To remember dialog save to path


Private Sub cmdExtract_Click()
Dim iFind As Long
Dim sExt As String
Dim sFile As String
Dim ifree As Integer
Dim iTerminator(1) As Long  'No need for more vars,
                            'so i used iTerminator
                            'as a array, my preference...


    On Error GoTo ExitSearch
    
    ifree = FreeFile
    Open txtPath For Binary Access Read As #ifree                           'Opens the file for reading
        sFile = Space(LOF(ifree))                                           'Puts the null terminator at the end of string variable
        Get #ifree, , sFile                                                 'Can now put the file into the variable cause it has space to accommodate it
    Close #ifree                                                            'Close the process
    
    sFile = LCase(sFile)                                                    'To prevent search ambiguity, make it non case sensitive
    sExt = LCase(txtExt)                                                    'Same thing
    
    Do
        iFind = InStr(iFind + 1, sFile, sExt)                               'Find the file extention in the string
        If iFind = 0 Then Exit Do
        iTerminator(0) = InStrRev(sFile, Chr(0), iFind)                     'Chr(0) is used to determinate the beginning of the file found and the ending
        'iTerminator(0) = pInstrRev(sFile, Chr(0), iFind)                   'For VB5 Users, comment the one above and uncomment this line
        iTerminator(1) = iFind + 4                                          'To determine the ending, ifind is the start of say for example ".dll", +4 cause theres 4 letters so it gets the ending
        If iTerminator(0) And Mid(sFile, iTerminator(1), 1) = Chr(0) Then   'Beginning point and end point of file in string
                If iTerminator(1) - iTerminator(0) - 1 < 20 _
                And iTerminator(1) - iTerminator(0) - 1 > 5 Then            'Some parameters to make sure the findings are not something other than files
                                                                            'This assumes all dlls are less than a length of 20 and greater than 1 char length.
                    lstFile.AddItem Mid(sFile, iTerminator(0) + 1, iTerminator(1) - iTerminator(0) - 1)
                End If
        End If
    Loop
    
    If chkDupes.Value Then KillDupesAPI lstFile                             'Remove doubles if checkbox is checked
    If lstFile.ListCount > 0 Then lstFile.ListIndex = 0                     'Set listindex to 0 in listbox
    
    Exit Sub
ExitSearch:
    MsgBox "Error#" & Err & " - " & Error(Err), vbCritical, "Error Occured"
End Sub

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
    SaveListBox lstFile, dFileName, chkAppend.Value
    Exit Sub
DialogError:
    MsgBox "Error#" & Err & " - " & Error(Err), vbCritical, "Error Occured"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.Width < 6100 Then Me.Width = 6100
    If Me.Height < 2265 Then Me.Height = 2265
    lstFile.Width = Me.Width - 360
    lstFile.Height = Me.Height - 2300
    frmOpt.Width = lstFile.Width
    txtPath.Width = frmOpt.Width - 920
    txtExt.Left = txtPath.Width + 120
    lblExt.Left = txtExt.Left - lblExt.Width - 150
    cmdBrowse.Left = txtPath.Width - 90
    cmdExtract.Left = cmdBrowse.Left - cmdExtract.Width - 100
    cmdClear.Left = cmdExtract.Left - cmdClear.Width - 100
    cmdSave.Left = cmdClear.Left - cmdSave.Width - 100
    'Some form resizing routines, nothing much
End Sub

Private Sub lstFile_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
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

Private Sub txtPath_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    txtPath = Data.Files(1)
    'Drag and drop a file into the txtpath box instead of using browse feature, your choice
End Sub
