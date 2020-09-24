VERSION 5.00
Object = "{19B7F2A2-1610-11D3-BF30-1AF820524153}#1.2#0"; "ccrpftv6.ocx"
Begin VB.Form frmPath 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Select Where to Start ...."
   ClientHeight    =   3420
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   4140
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cdmCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   120
      TabIndex        =   2
      Top             =   3015
      Width           =   1455
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "&Select"
      Height          =   330
      Left            =   2595
      TabIndex        =   1
      Top             =   3015
      Width           =   1455
   End
   Begin CCRPFolderTV6.FolderTreeview DIR 
      Height          =   2460
      Left            =   120
      TabIndex        =   0
      Top             =   420
      Width           =   3930
      _ExtentX        =   6932
      _ExtentY        =   4339
      RootFolder      =   "C:\"
      SelectedFolder  =   "C:\"
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "To change root folder visit settings tab."
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   105
      Width           =   2745
   End
End
Attribute VB_Name = "frmPath"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'                                                                    '
' Release Date: July 31, 2002                                        '
' Copyright © 2002 http://www.Europeum.net/, Vladimir S. Pekulas     '
'                                                                    '
' Search and Replace is a search utility that can find and replace   '
' one or more strings in multiple files. This application is         '
' released under GPL v.2 or higher.                                  '
'                                                                    '
' If you have any questions please feel free to let me know:         '
' vpekulas@europeum.net                                              '
'                                                                    '
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Option Explicit

'// UNLOAD FORM
Private Sub cdmCancel_Click()
    Unload Me
End Sub

'// GET SELECTED PATH AND APPLY IT TO DESIRED CONTROL
Private Sub cmdSelect_Click()
Dim gPATH As String
    gPATH = DIR.SelectedFolder
    If Not Mid(gPATH, Len(gPATH), 1) = "\" Then gPATH = gPATH & "\"

   If USE_TAB = 1 Then frmMain.coPath = gPATH
   If USE_TAB = 2 Then frmMain.txtPath = DIR.SelectedFolder
   Unload Me
End Sub

'// GET LAST KNOWN POSITION OF WINDOW AND APPLY ROOT PATH
Private Sub Form_Load()
On Error Resume Next
    If Not Trim(ROOT_PATH) = "" And USE_TAB = 1 Then
        DIR.RootFolder = ROOT_PATH
    End If
    Me.Left = GetSetting("PathWin", "Settings", "MainLeft", 1000)
    Me.Top = GetSetting("PathWin", "Settings", "MainTop", 1000)
    Me.Width = GetSetting("PathWin", "Settings", "MainWidth", 6500)
    Me.Height = GetSetting("PathWin", "Settings", "MainHeight", 6500)
End Sub

'// RESIZE CONTROLS
Private Sub Form_Resize()
On Error Resume Next
    If Me.Width > 4260 Or Me.Width < 4260 Then Me.Width = 4260
    DIR.Height = Me.Height - 1290
    cdmCancel.Top = Me.Height - 765
    cmdSelect.Top = Me.Height - 765
End Sub

'// SAVE POSITION OF WINDOW
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        SaveSetting "PathWin", "Settings", "MainLeft", Me.Left
        SaveSetting "PathWin", "Settings", "MainTop", Me.Top
        SaveSetting "PathWin", "Settings", "MainWidth", Me.Width
        SaveSetting "PathWin", "Settings", "MainHeight", Me.Height
    End If
End Sub

'// ESC KEY = UNLOAD FORM
Private Sub cdmCancel_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
Private Sub cmdSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub DIR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
