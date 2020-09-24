VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmORGFAV 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Favorites"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   6480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   510
      Left            =   150
      ScaleHeight     =   510
      ScaleWidth      =   3480
      TabIndex        =   2
      Top             =   3315
      Width           =   3480
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6660
         Top             =   465
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmORGFAV.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmORGFAV.frx":0112
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmORGFAV.frx":0224
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.Toolbar tlbMAIN 
         Height          =   330
         Left            =   0
         TabIndex        =   3
         Top             =   15
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   582
         ButtonWidth     =   1826
         ButtonHeight    =   582
         ToolTips        =   0   'False
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Delete "
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Open "
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Cancel "
               ImageIndex      =   3
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Available Favorites"
      Height          =   3165
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   6225
      Begin MSComctlLib.ListView lstLIST 
         Height          =   2805
         Left            =   135
         TabIndex        =   1
         Top             =   240
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   4948
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Favorite"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Search  String"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Replace String"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Mask"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Path"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "ID"
            Object.Width           =   706
         EndProperty
      End
   End
End
Attribute VB_Name = "frmORGFAV"
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

'// LOAD AVAILABLE FAVORITES
Private Sub Form_Load()
On Error Resume Next
    Call LOAD_FAVORITES
End Sub

'// SET FOCUS TO FAVORITES LISTING
Private Sub Form_Paint()
    lstLIST.SetFocus
End Sub

'// LOAD SELECTED FAVORITE
Sub lstLIST_DblClick()
On Error Resume Next
     Call LOAD_ID_FAVORITES(Int(Trim(lstLIST.SelectedItem.ListSubItems(5).Text)))
     Unload Me
End Sub

'// EITHER UNLOAD FORM (27) OR OPEN SELECTED FAVORITE ON ENTER (13)
Private Sub lstLIST_KeyPress(KeyAscii As Integer)
On Error Resume Next
    If KeyAscii = 27 Then Unload Me
    If KeyAscii = 13 Then lstLIST_DblClick
End Sub

'// MENU BUTTONS SELECTION (OPEN/DELETE/CANCEL)
Private Sub tlbMAIN_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
    Select Case Button.Index
            Case 3 ' CANCEL
                Unload Me
            Case 2 ' OPEN
                Call LOAD_ID_FAVORITES(Int(Trim(lstLIST.SelectedItem.ListSubItems(5).Text)))
                Unload Me
            Case 1 ' Delete
              If SHOW_VER = 1 Then
               If MsgBox("Delete favorite '" & Trim(lstLIST.SelectedItem.Text) & "' ?", vbOKCancel, "Delete Favorite ...") = vbOK Then
                    Call DELETE_FOVORITE(Int(Trim(lstLIST.SelectedItem.ListSubItems(5).Text)))
                    lstLIST.ListItems.Clear
                    Call EMPTY_FAV_MENU
                    Call LOAD_FAVORITES_MENU
                    Call LOAD_FAVORITES
               End If
              Else
                    Call DELETE_FOVORITE(Int(Trim(lstLIST.SelectedItem.ListSubItems(5).Text)))
                    lstLIST.ListItems.Clear
                    Call EMPTY_FAV_MENU
                    Call LOAD_FAVORITES_MENU
                    Call LOAD_FAVORITES
              End If
    End Select
End Sub
