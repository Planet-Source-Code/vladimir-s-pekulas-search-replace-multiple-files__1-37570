VERSION 5.00
Begin VB.Form frMAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About ..."
   ClientHeight    =   4725
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4095
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4725
   ScaleWidth      =   4095
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCopy 
      Height          =   1860
      Left            =   285
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1770
      Width           =   3495
   End
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&OK"
      Height          =   330
      Left            =   1965
      TabIndex        =   4
      Top             =   4275
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   40
      Left            =   285
      TabIndex        =   3
      Top             =   4095
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   975
      Left            =   285
      Picture         =   "frMAbout.frx":0000
      ScaleHeight     =   945
      ScaleWidth      =   3465
      TabIndex        =   0
      Top             =   150
      Width           =   3495
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.Europeum.net"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   0
      Left            =   285
      MouseIcon       =   "frMAbout.frx":AB8C
      MousePointer    =   99  'Custom
      TabIndex        =   5
      Top             =   3780
      Width           =   1860
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version: 1.00.1"
      Height          =   195
      Index           =   1
      Left            =   285
      TabIndex        =   2
      Top             =   1440
      Width           =   1065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Search And Replace by Europeum.net"
      Height          =   195
      Index           =   0
      Left            =   285
      TabIndex        =   1
      Top             =   1215
      Width           =   2730
   End
End
Attribute VB_Name = "frMAbout"
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

'// OK BUTTON
Private Sub cmdProcess_Click()
    Unload Me
End Sub

'// LOAD VERSION AND SET COPYRIGHT
Private Sub Form_Load()
    lblVersion(1).Caption = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
    txtCopy.Text = "You can redistribute this software under the terms of the GNU General Public License as published by the Free Software Foundation; either version 2. This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU General Public License for more details."
End Sub

'// OVERMOUSE LINK EFFECT
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblVersion(0).Font.Underline = True
    lblVersion(0).ForeColor = &H800000
End Sub

'// START AT TEXTBOX
Private Sub Form_Paint()
    txtCopy.SetFocus
End Sub

'// OPEN BROWSER WITH DEVELOPER'S URL
Private Sub lblVersion_Click(Index As Integer)
On Error Resume Next
    Shell ("start http://www.Europeum.net"), vbHide
End Sub

'// OVERMOUSE LINK EFFECT
Private Sub lblVersion_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblVersion(0).Font.Underline = False
    lblVersion(0).ForeColor = vbBlue
End Sub

'// UNLOAD ON ESC
Private Sub cmdProcess_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub txtCopy_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
