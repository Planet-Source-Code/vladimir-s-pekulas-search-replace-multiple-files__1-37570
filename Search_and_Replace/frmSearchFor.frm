VERSION 5.00
Begin VB.Form frmSearchFor 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Search For Block ..."
   ClientHeight    =   3465
   ClientLeft      =   60
   ClientTop       =   300
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3465
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   2940
      Left            =   90
      TabIndex        =   1
      Top             =   45
      Width           =   3390
      Begin VB.TextBox txtSearch 
         Height          =   2400
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   3
         Top             =   405
         Width           =   3165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Search For Block:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   135
         Width           =   1275
      End
   End
   Begin VB.CommandButton cdmCancel 
      Caption         =   "&OK"
      Height          =   330
      Left            =   2010
      TabIndex        =   0
      Top             =   3060
      Width           =   1455
   End
End
Attribute VB_Name = "frmSearchFor"
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

'// LOAD LAST KNOWN POSITION OF WINDOW
Private Sub Form_Load()
    Me.Left = GetSetting("SearchForWin", "Settings", "MainLeft", 1000)
    Me.Top = GetSetting("SearchForWin", "Settings", "MainTop", 1000)
    Me.Width = GetSetting("SearchForWin", "Settings", "MainWidth", 6500)
    Me.Height = GetSetting("SearchForWin", "Settings", "MainHeight", 6500)
End Sub

'// SAVE POSITION OF WINDOWS
Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting "SearchForWin", "Settings", "MainLeft", Me.Left
        SaveSetting "SearchForWin", "Settings", "MainTop", Me.Top
        SaveSetting "SearchForWin", "Settings", "MainWidth", Me.Width
        SaveSetting "SearchForWin", "Settings", "MainHeight", Me.Height
    End If
End Sub

'// APPLY TEXT TO DROP DOWN ON frmMAIN
Private Sub cdmCancel_Click()
    If Not Trim(txtSearch.Text) = "" Then
        Call SaveFile(App.PATH & "\srch.dat", txtSearch.Text & "sar|~|~|~|sar")
    End If
    frmMain.coSearch.Text = txtSearch.Text
    Unload Me
    BLOCK_SEARCH = True
End Sub

'// SET FOCUS
Private Sub Form_Paint()
    txtSearch.SetFocus
End Sub

'// RESIZE CONTROLS
Private Sub Form_Resize()
On Error Resume Next
 If Not Me.WindowState = 1 Then
    If Me.Width < 3675 Then Me.Width = 3675
    If Me.Height < 3825 Then Me.Height = 3825
    Frame1.Width = Me.Width - 260
    cdmCancel.Left = Me.Width - 1665
    Frame1.Height = Me.Height - 850
    cdmCancel.Top = Me.Height - 750
    txtSearch.Width = Frame1.Width - 190
    txtSearch.Height = Frame1.Height - 500
 End If
End Sub

'// ESC KEY = UNLOAD FORM
Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub cdmCancel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub

