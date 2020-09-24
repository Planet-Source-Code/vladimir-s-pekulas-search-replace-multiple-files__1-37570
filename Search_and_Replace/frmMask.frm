VERSION 5.00
Begin VB.Form frmMask 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mask Editor ..."
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   3555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cdmCancel 
      Caption         =   "&OK"
      Height          =   330
      Left            =   1980
      TabIndex        =   7
      Top             =   3105
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   2940
      Left            =   60
      TabIndex        =   0
      Top             =   90
      Width           =   3390
      Begin VB.CommandButton cmdRemove 
         Height          =   285
         Left            =   2745
         Picture         =   "frmMask.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Remove Extention"
         Top             =   630
         Width           =   510
      End
      Begin VB.CommandButton cmdAdd 
         Height          =   285
         Left            =   2745
         Picture         =   "frmMask.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   4
         ToolTipText     =   "Add Extention"
         Top             =   225
         Width           =   510
      End
      Begin VB.ListBox lstExt 
         Height          =   1815
         Left            =   135
         TabIndex        =   3
         Top             =   990
         Width           =   3120
      End
      Begin VB.TextBox txtEXT 
         Height          =   285
         Left            =   1710
         TabIndex        =   2
         Text            =   "*."
         Top             =   225
         Width           =   915
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Available Extentions:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   6
         Top             =   720
         Width           =   1470
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Add New Extention:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   255
         Width           =   1410
      End
   End
End
Attribute VB_Name = "frmMask"
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

'// UNLOAD WINDOW AND APPLY NEW FILE MASK TO DROP DOWN ON frmMAIN
Private Sub cdmCancel_Click()
Dim CONTENT As String, I As Integer
    For I = 0 To lstExt.ListCount
        CONTENT = CONTENT & lstExt.List(I) & "|"
    Next I
    Call SaveFile(App.PATH & "\ext.dat", CONTENT)
    Call Load_Masks
    Unload Me
End Sub

'// ADD NEW FILE MASK
Private Sub cmdAdd_Click()
    If Check_Ext(txtEXT.Text) = True Then
        lstExt.AddItem Replace(Trim(txtEXT.Text), " ", "")
        txtEXT.Text = "*."
        txtEXT.SelStart = 2
        txtEXT.SetFocus
    Else
        MsgBox "Please enter valid extention. Example:" & vbCrLf & "*.html,*.htm" & vbCrLf & "*.txt", vbInformation, "Valid extention ..."
    End If
End Sub

'// REMOVE SELECTED FILE MASK
Private Sub cmdRemove_Click()
    If Not lstExt.ListIndex = -1 Then lstExt.RemoveItem (lstExt.ListIndex)
End Sub

'// LOAD AVAILABLE FILE MASKS
Private Sub Form_Load()
On Error Resume Next
Dim I As Integer, FILE_EXT As String, UNQ_EXT_LN As Variant
    FILE_EXT = Open_File(App.PATH & "\ext.dat", False)
    UNQ_EXT_LN = Split(FILE_EXT, "|")
    For I = 0 To UBound(UNQ_EXT_LN)
        If Not Trim(UNQ_EXT_LN(I)) = "" Then lstExt.AddItem UNQ_EXT_LN(I)
    Next I
End Sub

'// CHECK THAT THE ENTERED EXTENTION IS IN FORMAT *.AAA+
Function Check_Ext(ByRef EXT As String) As Boolean
    If Len(Trim(EXT)) >= 2 And Mid(Trim(EXT), 1, 2) = "*." And InStr(Trim(EXT), "*.") = 1 And Not Len(EXT) <= 2 And InStr(Trim(EXT), "|") = 0 Then
        Check_Ext = True
    Else
        Check_Ext = False
    End If
End Function

'// ESC KEY = UNLOAD ME
Private Sub lstExt_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
Private Sub txtEXT_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
Private Sub cmdRemove_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
Private Sub cmdAdd_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
Private Sub cdmCancel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then Unload Me
End Sub
