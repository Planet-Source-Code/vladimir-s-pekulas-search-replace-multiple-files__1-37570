VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmEdit 
   Caption         =   "File Content ..."
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8640
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   8640
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox pHolder 
      BorderStyle     =   0  'None
      Height          =   465
      Left            =   120
      ScaleHeight     =   465
      ScaleWidth      =   4800
      TabIndex        =   1
      Top             =   6360
      Width           =   4800
      Begin MSComctlLib.Toolbar tlbGeneral 
         Height          =   330
         Left            =   75
         TabIndex        =   2
         Top             =   60
         Width           =   4020
         _ExtentX        =   7091
         _ExtentY        =   582
         ButtonWidth     =   2249
         ButtonHeight    =   582
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Save As ..."
               Object.ToolTipText     =   "Start listing available files and it's information."
               ImageIndex      =   2
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Print"
               Object.ToolTipText     =   "Process checked files."
               ImageIndex      =   6
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "&Cancel  "
               ImageIndex      =   7
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fmMain 
      Caption         =   "File Content"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7560
         Top             =   360
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEdit.frx":0442
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEdit.frx":0894
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEdit.frx":09A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEdit.frx":0DF8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEdit.frx":124A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEdit.frx":135C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmEdit.frx":146E
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin RichTextLib.RichTextBox RTF 
         Height          =   5775
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   10186
         _Version        =   393217
         Enabled         =   -1  'True
         ScrollBars      =   3
         TextRTF         =   $"frmEdit.frx":1580
      End
   End
End
Attribute VB_Name = "frmEdit"
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

'// LOAD LAST KNOWN POSITION OF THE WINDOW
Private Sub Form_Load()
    Me.Left = GetSetting(App.Title, "Content", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Content", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Content", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Content", "MainHeight", 6500)
    RTF.RightMargin = RTF.Width * 50
    RTF.LoadFile FILE_TO_OPEN
End Sub

'// RESIZE CONTROLS
Private Sub Form_Resize()
On Error Resume Next
 If Not Me.WindowState = 1 Then
    If Me.Width < 5160 Then Me.Width = 5160
    If Me.Height < 3000 Then Me.Height = 3000
    fmMain.Width = Me.Width - 400
    fmMain.Height = Me.Height - 1100
    RTF.Width = fmMain.Width - 235
    RTF.Height = fmMain.Height - 350
    pHolder.Top = Me.Height - 900
 End If
End Sub

'// SAVE POSITION FOR FUTURE USE
Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Content", "MainLeft", Me.Left
        SaveSetting App.Title, "Content", "MainTop", Me.Top
        SaveSetting App.Title, "Content", "MainWidth", Me.Width
        SaveSetting App.Title, "Content", "MainHeight", Me.Height
    End If
End Sub

'// MENU BUTTON SELECTION (SAVE AS/PRINT/CANCEL)
Private Sub tlbGeneral_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1
            With frmMain.CDC
                .DialogTitle = "Save file as ..."
                 If AS_HTML = 1 Then
                 .DefaultExt = "html"
                 Else
                 .DefaultExt = "txt"
                 End If
                 .FileName = FILE_TO_OPEN
                 .CancelError = False
                .ShowSave
                If Not Trim(.FileName) = "" Then Call SaveFile(.FileName, RTF.Text)
            End With
        Case 2
            On Error Resume Next
            With frmMain.CDC
                .DialogTitle = "Print Report Log ..."
                .CancelError = False
                .Flags = cdlPDReturnDC + cdlPDNoPageNums
                If RTF.SelLength = 0 Then
                    .Flags = .Flags + cdlPDAllPages
                Else
                    .Flags = .Flags + cdlPDSelection
                End If
                .ShowPrinter
                If ERR <> MSComDlg.cdlCancel Then
                    RTF.SelPrint .hDC
                End If
            End With
        Case 3
            Unload Me
    End Select
End Sub
