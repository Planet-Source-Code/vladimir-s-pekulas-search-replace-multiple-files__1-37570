VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   2340
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2340
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1500
      Left            =   4440
      Top             =   75
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   30
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   945
      ScaleWidth      =   3465
      TabIndex        =   0
      Top             =   30
      Width           =   3465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright © 2002, Released under GPL v.2, Vladimir S. Pekulas"
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   2055
      Width           =   4530
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   405
      Left            =   -105
      Top             =   1935
      Width           =   6000
   End
End
Attribute VB_Name = "frmSplash"
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

'// LOAD frmMAIN INTO MEMORY, AND CHECK IF WE SHOULD SHOW SPLASH SCREEN
Private Sub Form_Load()
    Load frmMain
    
    If Not SHOW_SPLASH = 1 Then
        frmMain.Show
        Unload Me
    End If
End Sub

'// KEEP SHOWING THE FORM IF DESIRED FOR THE LENGTH OF TIMER VALUE
Private Sub Timer1_Timer()
    frmMain.Show
    Unload Me
End Sub
