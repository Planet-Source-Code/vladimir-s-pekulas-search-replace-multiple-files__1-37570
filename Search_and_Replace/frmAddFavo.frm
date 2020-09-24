VERSION 5.00
Begin VB.Form frmAddFavo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add to favorites ..."
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   4500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdProcess 
      Caption         =   "&ADD"
      Height          =   330
      Left            =   2565
      TabIndex        =   2
      Top             =   900
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Favorite Name: "
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   75
      Width           =   4290
      Begin VB.TextBox txtNAME 
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Top             =   270
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmAddFavo"
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

'// ADD FAVORITE
Private Sub cmdProcess_Click()
Dim strSETTINGS As String, strTOUCH As String, strCOLOR As String
    With frmMain
        strSETTINGS = .chCase.Value & "," & .chFind.Value & "," & .chKeep.Value & "," & .chSub.Value & "," & .chVer.Value & "," & .chWords.Value & "," & .chMulti.Value & "," & .chTouch.Value & "," & .chSM.Value
        strTOUCH = .chA.Value & "," & .chR.Value & "," & .chH.Value
        strCOLOR = .pCOLOR.BackColor & "," & .pCOLORun.BackColor
    End With
    Call ADD_FOVORITE(Trim(txtNAME.Text), frmMain.coSearch.Text, frmMain.coReplace.Text, frmMain.coMask.Text, frmMain.coPath.Text, strSETTINGS, strTOUCH, strCOLOR, frmMain.txtPath.Text)
    Call EMPTY_FAV_MENU
    Call LOAD_FAVORITES_MENU
    Unload Me
End Sub

'// ESC & ENTER KEY
Private Sub txtNAME_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then Unload Me
  If KeyAscii = 13 Then cmdProcess.Value = True
End Sub
Private Sub cmdProcess_KeyPress(KeyAscii As Integer)
  If KeyAscii = 27 Then Unload Me
End Sub
