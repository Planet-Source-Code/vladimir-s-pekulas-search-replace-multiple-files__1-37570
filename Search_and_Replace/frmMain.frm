VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmMain 
   Caption         =   "Search and Replace by Europeum.net"
   ClientHeight    =   6660
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   7965
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   7965
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   75
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   11456
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " General "
      TabPicture(0)   =   "frmMain.frx":0442
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblTitle"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "coSearch"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "coReplace"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "coMask"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "coPath"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdOpen"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdMask"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Command2"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Command3"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lstFiles"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "cmdCheckAll"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "chAll"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "ImageList1"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Dir1"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "File1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "chUNCH"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "pHolder3"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "rtfHIDE"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Settings"
      TabPicture(1)   =   "frmMain.frx":04A0
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label7"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "chCase"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "chFind"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chKeep"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "optHTML"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "optText"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtPath"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdOpenPath"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "chSub"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "chVer"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "chWords"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "chMulti"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "chTouch"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "frTouch"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "chSM"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cmdColor"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "pCOLOR"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "pCOLORun"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmdColor2"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "pHolder1"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "chSplash"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).ControlCount=   23
      TabCaption(2)   =   " Report"
      TabPicture(2)   =   "frmMain.frx":051B
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "picHolder"
      Tab(2).Control(1)=   "CDC"
      Tab(2).Control(2)=   "RTF_LOG"
      Tab(2).Control(3)=   "Label5"
      Tab(2).ControlCount=   4
      Begin RichTextLib.RichTextBox rtfHIDE 
         Height          =   285
         Left            =   6930
         TabIndex        =   52
         Top             =   15000
         Visible         =   0   'False
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   503
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmMain.frx":06FB
      End
      Begin VB.CheckBox chSplash 
         Caption         =   "Show Splash Screen"
         Height          =   195
         Left            =   -74640
         TabIndex        =   51
         ToolTipText     =   "Check if you want to see the splash screen on strat-up"
         Top             =   3510
         Width           =   2085
      End
      Begin VB.PictureBox pHolder3 
         BorderStyle     =   0  'None
         Height          =   465
         Left            =   225
         ScaleHeight     =   465
         ScaleWidth      =   5280
         TabIndex        =   49
         Top             =   5895
         Width           =   5280
         Begin MSComctlLib.Toolbar tlbGeneral 
            Height          =   330
            Left            =   75
            TabIndex        =   50
            Top             =   60
            Width           =   4125
            _ExtentX        =   7276
            _ExtentY        =   582
            ButtonWidth     =   3572
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Appearance      =   1
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "&Process  "
                  Object.ToolTipText     =   "Start listing available files and it's information."
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Enabled         =   0   'False
                  Caption         =   "&Search And Replace "
                  Object.ToolTipText     =   "Process checked files."
                  ImageIndex      =   5
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox pHolder1 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -74775
         ScaleHeight     =   375
         ScaleWidth      =   3300
         TabIndex        =   47
         Top             =   5760
         Width           =   3300
         Begin MSComctlLib.Toolbar tlbSettings 
            Height          =   330
            Left            =   30
            TabIndex        =   48
            Top             =   15
            Width           =   2360
            _ExtentX        =   4154
            _ExtentY        =   582
            ButtonWidth     =   2011
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Appearance      =   1
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "&Save   "
                  Object.ToolTipText     =   "Saves selected settings"
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "&Default   "
                  Object.ToolTipText     =   "Assigns default values to all settings"
                  ImageIndex      =   4
               EndProperty
            EndProperty
         End
      End
      Begin VB.PictureBox picHolder 
         BorderStyle     =   0  'None
         Height          =   375
         Left            =   -74775
         ScaleHeight     =   375
         ScaleWidth      =   3705
         TabIndex        =   45
         Top             =   5800
         Width           =   3705
         Begin MSComctlLib.Toolbar Toolbar1 
            Height          =   330
            Left            =   60
            TabIndex        =   46
            Top             =   15
            Width           =   3120
            _ExtentX        =   5503
            _ExtentY        =   582
            ButtonWidth     =   2381
            ButtonHeight    =   582
            AllowCustomize  =   0   'False
            Wrappable       =   0   'False
            Appearance      =   1
            Style           =   1
            TextAlignment   =   1
            ImageList       =   "ImageList1"
            _Version        =   393216
            BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
               NumButtons      =   2
               BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   "&Save as ...  "
                  Object.ToolTipText     =   "Save generated log as ..."
                  ImageIndex      =   2
               EndProperty
               BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
                  Caption         =   " &Print     "
                  Object.ToolTipText     =   "Print generated log ..."
                  ImageIndex      =   6
               EndProperty
            EndProperty
         End
      End
      Begin VB.CommandButton chUNCH 
         Height          =   315
         Left            =   6885
         Picture         =   "frmMain.frx":077D
         Style           =   1  'Graphical
         TabIndex        =   9
         ToolTipText     =   "Check only files that match"
         Top             =   2490
         Width           =   420
      End
      Begin VB.CommandButton cmdColor2 
         Caption         =   "..."
         Height          =   285
         Left            =   -68925
         TabIndex        =   27
         Top             =   4275
         Width           =   420
      End
      Begin VB.PictureBox pCOLORun 
         BackColor       =   &H8000000A&
         Height          =   285
         Left            =   -72390
         ScaleHeight     =   225
         ScaleWidth      =   3285
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Selected un-matched file color."
         Top             =   4275
         Width           =   3345
      End
      Begin VB.PictureBox pCOLOR 
         Height          =   285
         Left            =   -72390
         ScaleHeight     =   225
         ScaleWidth      =   3285
         TabIndex        =   41
         TabStop         =   0   'False
         ToolTipText     =   "Selected matched file color."
         Top             =   3900
         Width           =   3345
      End
      Begin VB.CommandButton cmdColor 
         Caption         =   "..."
         Height          =   285
         Left            =   -68925
         TabIndex        =   26
         Top             =   3900
         Width           =   420
      End
      Begin VB.CheckBox chSM 
         Caption         =   "Show Matches (Slower)"
         Height          =   195
         Left            =   -74640
         TabIndex        =   25
         ToolTipText     =   "Check to see the number of matches on file listing"
         Top             =   3195
         Width           =   2220
      End
      Begin VB.Frame frTouch 
         Caption         =   "Touch Only Files with Attributes: "
         Height          =   735
         Left            =   -72390
         TabIndex        =   40
         Top             =   2670
         Width           =   3390
         Begin VB.CheckBox chR 
            Caption         =   "Read Only"
            Height          =   195
            Left            =   1170
            TabIndex        =   23
            Top             =   315
            Width           =   1275
         End
         Begin VB.CheckBox chH 
            Caption         =   "Hidden"
            Height          =   195
            Left            =   2430
            TabIndex        =   24
            Top             =   315
            Width           =   915
         End
         Begin VB.CheckBox chA 
            Caption         =   "Achive"
            Height          =   195
            Left            =   180
            TabIndex        =   22
            Top             =   315
            Width           =   1005
         End
      End
      Begin VB.CheckBox chTouch 
         Caption         =   "Apply Touch"
         Height          =   195
         Left            =   -74640
         TabIndex        =   21
         ToolTipText     =   "Check to touch only files with selected attributes."
         Top             =   2895
         Width           =   2220
      End
      Begin MSComDlg.CommonDialog CDC 
         Left            =   -68205
         Top             =   5085
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CheckBox chMulti 
         Caption         =   "Multi Line Search"
         Height          =   195
         Left            =   -74640
         TabIndex        =   20
         ToolTipText     =   "To preform block search and replace, this value must be checked."
         Top             =   2595
         Width           =   2220
      End
      Begin VB.CheckBox chWords 
         Caption         =   "Whole Words Only"
         Height          =   240
         Left            =   -74640
         TabIndex        =   19
         ToolTipText     =   "Check this value to search and replace only whole words, rather then portion of a string."
         Top             =   2250
         Width           =   3030
      End
      Begin RichTextLib.RichTextBox RTF_LOG 
         Height          =   4740
         Left            =   -74775
         TabIndex        =   30
         Top             =   945
         Width           =   7215
         _ExtentX        =   12726
         _ExtentY        =   8361
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmMain.frx":0C89
      End
      Begin VB.CheckBox chVer 
         Caption         =   "Show Verification"
         Height          =   240
         Left            =   -74640
         TabIndex        =   18
         ToolTipText     =   "Check this value to see a verification before proceeding to selected action."
         Top             =   1935
         Width           =   1770
      End
      Begin VB.CheckBox chSub 
         Caption         =   "Include Sub Folders"
         Height          =   240
         Left            =   -74640
         TabIndex        =   17
         ToolTipText     =   "Check this value to include files located in sub folders in your search and replace."
         Top             =   1620
         Width           =   1860
      End
      Begin VB.FileListBox File1 
         Enabled         =   0   'False
         Height          =   675
         Left            =   6030
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   25000
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.DirListBox Dir1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6030
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   25000
         Visible         =   0   'False
         Width           =   600
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6075
         Top             =   5130
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0D0B
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":115D
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":126F
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":16C1
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1B13
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1C25
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdOpenPath 
         Height          =   315
         Left            =   -68925
         Picture         =   "frmMain.frx":1D37
         Style           =   1  'Graphical
         TabIndex        =   28
         Top             =   4890
         Width           =   420
      End
      Begin VB.TextBox txtPath 
         Height          =   285
         Left            =   -72390
         TabIndex        =   29
         ToolTipText     =   "Path where the folder browser starts."
         Top             =   4905
         Width           =   3345
      End
      Begin VB.CheckBox chAll 
         Height          =   240
         Left            =   6885
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   3375
         Width           =   285
      End
      Begin VB.CommandButton cmdCheckAll 
         Height          =   315
         Left            =   6885
         Picture         =   "frmMain.frx":1E39
         Style           =   1  'Graphical
         TabIndex        =   10
         ToolTipText     =   "Check/Uncheck listed files"
         Top             =   2940
         Width           =   420
      End
      Begin VB.OptionButton optText 
         Caption         =   "Text"
         Height          =   195
         Left            =   -71310
         TabIndex        =   16
         Top             =   1305
         Width           =   1005
      End
      Begin VB.OptionButton optHTML 
         Caption         =   "HTML"
         Height          =   195
         Left            =   -72390
         TabIndex        =   15
         Top             =   1305
         Width           =   1005
      End
      Begin VB.CheckBox chKeep 
         Caption         =   "Keep Detailed Log"
         Height          =   195
         Left            =   -74640
         TabIndex        =   14
         ToolTipText     =   "Check this value to keep a detailed log of preformed search and replace."
         Top             =   1305
         Width           =   2040
      End
      Begin VB.CheckBox chFind 
         Caption         =   "Find All Occurrences"
         Height          =   195
         Left            =   -74640
         TabIndex        =   13
         ToolTipText     =   "Check this value to search and replace all occurrences of search string, rather then only the first one."
         Top             =   990
         Width           =   2760
      End
      Begin VB.CheckBox chCase 
         Caption         =   "Case Sensitive"
         Height          =   195
         Left            =   -74640
         TabIndex        =   12
         ToolTipText     =   "Check this value to respect the case of search string."
         Top             =   690
         Width           =   2580
      End
      Begin MSComctlLib.ListView lstFiles 
         Height          =   3300
         Left            =   315
         TabIndex        =   11
         Top             =   2490
         Width           =   6450
         _ExtentX        =   11377
         _ExtentY        =   5821
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Name"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Path"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Type"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Matches"
            Object.Width           =   1411
         EndProperty
      End
      Begin VB.CommandButton Command3 
         Height          =   315
         Left            =   6885
         Picture         =   "frmMain.frx":2345
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   630
         Width           =   420
      End
      Begin VB.CommandButton Command2 
         Height          =   315
         Left            =   6885
         Picture         =   "frmMain.frx":2447
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1005
         Width           =   420
      End
      Begin VB.CommandButton cmdMask 
         Height          =   315
         Left            =   6885
         Picture         =   "frmMain.frx":2549
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1395
         Width           =   420
      End
      Begin VB.CommandButton cmdOpen 
         Height          =   315
         Left            =   6885
         Picture         =   "frmMain.frx":264B
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1770
         Width           =   420
      End
      Begin VB.ComboBox coPath 
         Height          =   315
         Left            =   1395
         TabIndex        =   7
         ToolTipText     =   "Where to start looking"
         Top             =   1770
         Width           =   5370
      End
      Begin VB.ComboBox coMask 
         Height          =   315
         Left            =   1395
         TabIndex        =   5
         ToolTipText     =   "Example: *.txt,*.*asp"
         Top             =   1395
         Width           =   5370
      End
      Begin VB.ComboBox coReplace 
         Height          =   315
         Left            =   1395
         TabIndex        =   3
         ToolTipText     =   "Replace found string with"
         Top             =   1005
         Width           =   5370
      End
      Begin VB.ComboBox coSearch 
         Height          =   315
         Left            =   1395
         TabIndex        =   1
         ToolTipText     =   "Search for a string"
         Top             =   630
         Width           =   5370
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Un-Matched File Color:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   44
         Top             =   4320
         Width           =   1620
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Matched File Color:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   42
         Top             =   3945
         Width           =   1365
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Search and replace report log:"
         Height          =   195
         Left            =   -74775
         TabIndex        =   39
         Top             =   690
         Width           =   2145
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Root Folder for path selection:"
         Height          =   195
         Left            =   -74640
         TabIndex        =   36
         Top             =   4950
         Width           =   2130
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Path:"
         Height          =   195
         Left            =   315
         TabIndex        =   34
         Top             =   1830
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Replace With:"
         Height          =   195
         Left            =   315
         TabIndex        =   33
         Top             =   1065
         Width           =   1020
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File Mask:"
         Height          =   195
         Left            =   315
         TabIndex        =   32
         Top             =   1455
         Width           =   720
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "Search For:"
         Height          =   195
         Left            =   315
         TabIndex        =   31
         Top             =   690
         Width           =   825
      End
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Action"
      Begin VB.Menu mnuSettings 
         Caption         =   "&Settings"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuReport 
         Caption         =   "&View Report"
         Shortcut        =   {F3}
      End
      Begin VB.Menu mnusep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "&Export"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuImport 
         Caption         =   "&Import"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnusep7 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuCopy 
         Caption         =   "Copy Processed &Results"
         Shortcut        =   ^Q
      End
      Begin VB.Menu mnuPrinResults 
         Caption         =   "&Print Processed Results"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuViewRESHTML 
         Caption         =   "&View Processed Results as HTML"
      End
   End
   Begin VB.Menu mnuFavorites 
      Caption         =   "&Favorites"
      Begin VB.Menu mnuAddFavorite 
         Caption         =   "&Add to Favorite"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuOrganizeFav 
         Caption         =   "&Organize Favorites"
         Shortcut        =   {F11}
      End
      Begin VB.Menu mnusep6 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "A&bout"
      Begin VB.Menu mnuAboutAbout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnuContents 
         Caption         =   "&Help"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu mnuFileMenu 
      Caption         =   "FileMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuOpenAssoc 
         Caption         =   "&Open With Associated Program"
      End
      Begin VB.Menu mnuViewFile 
         Caption         =   "&View File Content"
      End
      Begin VB.Menu mnusep8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCheckUncheck 
         Caption         =   "&Check/Uncheck File"
      End
   End
End
Attribute VB_Name = "frmMain"
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
Const HH_DISPLAY_TOPIC = &H0 ' used for the CHM help file

'// ENABLE/DISABLE TOUCH SCREEN
Sub chTouch_Click()
    If chTouch.Value = 1 Then
        frTouch.Enabled = True
        chA.Enabled = True
        chH.Enabled = True
        chR.Enabled = True
     Else
        frTouch.Enabled = False
        chA.Enabled = False
        chH.Enabled = False
        chR.Enabled = False
     End If
End Sub

'// CHECK ONLY FILES THAT ARE A MATCH TO SEARCH STRING
Private Sub chUNCH_Click()
Dim I As Integer
If lstFiles.ListItems.Count = 0 Then Exit Sub
    For I = 1 To lstFiles.ListItems.Count
        If lstFiles.ListItems(I).ForeColor = NOTFOUND_CLR Then
            lstFiles.ListItems(I).Checked = False
        Else
            lstFiles.ListItems(I).Checked = True
        End If
    Next I
End Sub

'// SHOW COLOR DIALOG
Private Sub cmdColor_Click()
With CDC
    .DialogTitle = "Matched file color ..."
    .ShowColor
    pCOLOR.BackColor = .COLOR
    FOUND_CLR = .COLOR
End With
End Sub

'// SHOW COLOR DIALOG
Private Sub cmdColor2_Click()
With CDC
    .ShowColor
    pCOLORun.BackColor = .COLOR
    NOTFOUND_CLR = .COLOR
End With

End Sub

'// SHOW MASK DIALOG
Private Sub cmdMask_Click()
    frmMask.Show 1, Me
End Sub

'// OPEN PATH DIALOG
Private Sub cmdOpen_Click()
    USE_TAB = 1
    frmPath.Show 1, Me
End Sub

'// SET ROOT FOLDER
Private Sub cmdOpenPath_Click()
    USE_TAB = 2
    frmPath.Show 1, Me
End Sub

'// SHOW BLOCK REPLACE DIALOG
Private Sub Command2_Click()
    If Not MULTI_LINE = 1 Then
        MsgBox "For block search and replace Multi Line value must be checked." & vbCrLf & "Please do so under settings tab.", vbInformation, "Multi line ..."
        SSTab1.Tab = 1
        Exit Sub
    Else
        frmReplace.Show 1, Me
    End If
End Sub

'// SHOW SEARCH BLOCK DIALOG
Private Sub Command3_Click()
    If Not MULTI_LINE = 1 Then
        MsgBox "For block search and replace Multi Line value must be checked." & vbCrLf & "Please do so under settings tab.", vbInformation, "Multi line ..."
        SSTab1.Tab = 1
        Exit Sub
    Else
        frmSearchFor.Show 1, Me
    End If
End Sub

'// LIST FILES FOR FOLDER
Private Sub Dir1_Change()
    File1.PATH = Dir1.PATH
End Sub

'// LOAD SETTINGS, USED STRINGS AND FORM POSITION
Private Sub Form_Load()
    Call Load_Settings(True)
    Call Load_Masks
    Call Load_Search
    Call Load_Replace
    Call Load_Path
    Call LOAD_FAVORITES_MENU
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
End Sub

'// CHECK / UNCHECK ALL ITEMS IN LISTBOX
Private Sub cmdCheckAll_Click()
Dim I As Integer
    For I = 1 To lstFiles.ListItems.Count
        If chAll.Value = 1 Then lstFiles.ListItems(I).Checked = True
        If chAll.Value = 0 Then lstFiles.ListItems(I).Checked = False
    Next I
End Sub

'// THE VERIFICATION MESSAGE BEFORE PROCEEDING
Function Create_Message() As String
    Create_Message = ""
    Create_Message = Create_Message & "Search KeyWord : " & Trim(coSearch.Text) & vbCrLf
    Create_Message = Create_Message & "Replace KeyWord: " & Trim(coReplace.Text) & vbCrLf
    Create_Message = Create_Message & "Start Path     : " & Trim(coPath.Text) & vbCrLf & vbCrLf
    If CASE_SENS = 1 Then
    Create_Message = Create_Message & "Case Sensitive: YES" & vbCrLf
    Else
    Create_Message = Create_Message & "Case Sensitive: NO" & vbCrLf
    End If
    If SUB_DIR = 1 Then
    Create_Message = Create_Message & "Include Sub-Folders: YES" & vbCrLf
    Else
    Create_Message = Create_Message & "Include Sub-Folders: NO" & vbCrLf
    End If
    If FIND_ALL = 1 Then
    Create_Message = Create_Message & "Find All Occurrences: YES" & vbCrLf
    Else
    Create_Message = Create_Message & "Find All Occurrences: NO" & vbCrLf
    End If
    If WHOLE_W = 1 Then
    Create_Message = Create_Message & "Whole Words Only: YES" & vbCrLf
    Else
    Create_Message = Create_Message & "Whole Words Only: NO" & vbCrLf
    End If
    If MULTI_LINE = 1 Then
    Create_Message = Create_Message & "Multi Line: YES" & vbCrLf
    Else
    Create_Message = Create_Message & "Multi Line: NO" & vbCrLf
    End If
End Function

'// PROCESS FILES THAT ARE TO BE SEARCHED AND REPALCED
Sub Replace_Matches()
Dim I As Integer

    If AS_HTML = 0 Then '// TEXT VERSION
        REPORT_SECTION = ""
        REPORT_SECTION = REPORT_SECTION & "|-------------------------------------|" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "|-- >> SEARCH AND REPLACE REPORT << --|" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "|-------------------------------------|" & vbCrLf & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "Search  String : " & coSearch.Text & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "Replace String : " & coReplace.Text & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "File Mask      : " & coMask.Text & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "Selected Path  : " & coPath.Text & vbCrLf
        If chSub.Value = 1 Then
        REPORT_SECTION = REPORT_SECTION & "Sub-Folders    : True" & vbCrLf
        Else
        REPORT_SECTION = REPORT_SECTION & "Sub-Folders    : False" & vbCrLf
        End If
        If chCase.Value = 1 Then
        REPORT_SECTION = REPORT_SECTION & "Case Sensitive : True" & vbCrLf
        Else
        REPORT_SECTION = REPORT_SECTION & "Case Sensitive : False" & vbCrLf
        End If
        If chFind.Value = 1 Then
        REPORT_SECTION = REPORT_SECTION & "All Occurrences: True" & vbCrLf
        Else
        REPORT_SECTION = REPORT_SECTION & "All Occurrences: False" & vbCrLf
        End If
        If chTouch.Value = 1 Then
        REPORT_SECTION = REPORT_SECTION & "Applying Touch : True" & vbCrLf
        Else
        REPORT_SECTION = REPORT_SECTION & "Applying Touch : False" & vbCrLf & vbCrLf & vbCrLf
        End If
    Else            '//HTML VERSION
        REPORT_SECTION = ""
        REPORT_SECTION = REPORT_SECTION & "<!DOCTYPE HTML PUBLIC -//W3C//DTD HTML 4.01 Transitional//EN>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "<html><head><title>Search & Replace Report Log</title>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "<style type=text/css>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "  body,p,td{font-family:Verdana, Arial, Helvetica, sans-serif;font-size:13px;color: Black}" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "  a:link    {color:navy;text-decoration:underline}" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "  a:visited {color:navy;text-decoration:underline}" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "  a:hover   {color:blue;text-decoration:normal}" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "</style></head><body bgcolor=#ffffff>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "<table width=95% align=center cellpadding=2 cellspacing=0 border=1 bordercolor=WhiteSmoke><tr>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "<td><b>Search & Replace Report Log</b></td>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "</tr></table><br>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "<table width=95% align=center cellpadding=2 cellspacing=0 border=1 bordercolor=#ffffff><tr><td width=150>Search  String</td>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "<td bgcolor=WhiteSmoke>" & coSearch.Text & "</td>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "</tr><tr><td>Replace String</td>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "<td bgcolor=WhiteSmoke>" & coReplace.Text & "</td>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "</tr><tr><td>File Mask</td>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "<td bgcolor=WhiteSmoke>" & coMask.Text & "</td>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "</tr><tr><td>Selected Path</td>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "<td bgcolor=WhiteSmoke>" & coPath.Text & "</td>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "</tr><tr><td>Include Sub-Folders</td>" & vbCrLf
        If chSub.Value = 1 Then
        REPORT_SECTION = REPORT_SECTION & "<td bgcolor=WhiteSmoke>True</td>" & vbCrLf
        Else
        REPORT_SECTION = REPORT_SECTION & "<td bgcolor=WhiteSmoke>False</td>" & vbCrLf
        End If
        REPORT_SECTION = REPORT_SECTION & "</tr><tr><td>Case Sensitive</td>" & vbCrLf
        If chCase.Value = 1 Then
        REPORT_SECTION = REPORT_SECTION & "<td bgcolor=WhiteSmoke>True</td>" & vbCrLf
        Else
        REPORT_SECTION = REPORT_SECTION & "<td bgcolor=WhiteSmoke>False</td>" & vbCrLf
        End If
        REPORT_SECTION = REPORT_SECTION & "</tr><tr><td>All Occurrences</td>" & vbCrLf
        If chFind.Value = 1 Then
        REPORT_SECTION = REPORT_SECTION & "<td bgcolor=WhiteSmoke>True</td>" & vbCrLf
        Else
        REPORT_SECTION = REPORT_SECTION & "<td bgcolor=WhiteSmoke>False</td>" & vbCrLf
        End If
        REPORT_SECTION = REPORT_SECTION & "</tr><tr><td>Applying Touch:</td>" & vbCrLf
        If chTouch.Value = 1 Then
        REPORT_SECTION = REPORT_SECTION & "<td bgcolor=WhiteSmoke>True</td>" & vbCrLf
        Else
        REPORT_SECTION = REPORT_SECTION & "<td bgcolor=WhiteSmoke>False</td>" & vbCrLf
        End If
        
        
        REPORT_SECTION = REPORT_SECTION & "</tr></table><br><br>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "<table width=95% align=center cellpadding=2 cellspacing=0 border=0><tr><td valign=top>" & vbCrLf
    End If
    For I = 1 To lstFiles.ListItems.Count
        If lstFiles.ListItems(I).Checked = True Then '// ONLY CHECKED FILES
                Call Replace_Finall(lstFiles.ListItems(I).ListSubItems(1).Text & lstFiles.ListItems(I).Text, lstFiles.ListItems(I).ListSubItems(3).Text) '// 1 = Path 2 = Ext 3 = Matches
        End If
    Next I
    If AS_HTML = 1 Then '// HTML VERSION
        REPORT_SECTION = REPORT_SECTION & "</td></tr><tr><td bgcolor=WhiteSmoke><a href=http://www.europeum.net style=text-decoration:none;>Developed by Europeum.net</a></td></tr></table></body></html>" & vbCrLf
        frmMain.RTF_LOG.Text = REPORT_SECTION
    Else
        frmMain.RTF_LOG.Text = REPORT_SECTION
    End If
End Sub

'// REPLACE USING RE
Sub Replace_Finall(ByRef PATH As String, MTCH As String)
Dim STR_ORIGINAL As String
Dim re As New RegExp, ma As Match
    STR_ORIGINAL = Open_File(PATH, True)
    If MULTI_LINE = 1 Then
        re.MultiLine = True
    Else
        re.MultiLine = False
    End If
    If Not WHOLE_W = 1 Then                     '// ANY MATCH
        re.Pattern = frmMain.coSearch
    Else                                        '// WHOLE WORDS ONLY
        re.Pattern = "\b" & frmMain.coSearch & "\b"
    End If
    
    If CASE_SENS = 1 Then
        re.IgnoreCase = False    '//  case sensitive search \\
    Else                        '//
        re.IgnoreCase = True   '//
    End If                      '// ---- \\
    If FIND_ALL = 1 Then        '// find all the occurrences \\
        re.Global = True        '//
    Else                        '//
        re.Global = False       '//
    End If                      '// ---- \\
    Call SaveFile(PATH, re.Replace(STR_ORIGINAL, coReplace.Text))
    
    If AS_HTML = 0 Then '// TEXT VERSION
        REPORT_SECTION = REPORT_SECTION & "Processing File :   " & PATH & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "Found Matches   :   " & MTCH & vbCrLf
        REPORT_SECTION = REPORT_SECTION & ">>" & vbCrLf & vbCrLf
    Else
        REPORT_SECTION = REPORT_SECTION & "Processing File :   <a href=" & PATH & ">" & PATH & "</a><br>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "Found Matches   :   " & MTCH & "<br>" & vbCrLf
        REPORT_SECTION = REPORT_SECTION & "<hr style='height: 1px; color:silver;'><br>" & vbCrLf
    End If
End Sub

'// SAVE LAST POSITION OF THE FORM
Private Sub Form_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
End Sub

'// OPEN FILE ON DBL_CLICK
Private Sub lstFiles_DblClick()
On Error Resume Next
'// OPEN ON DBL_CLICK
    If Not lstFiles.ListItems.Count = 0 Then Shell ("start " & lstFiles.SelectedItem.ListSubItems(1).Text & "\" & lstFiles.SelectedItem.Text), vbHide
End Sub

'// SORT BY COLUMN
Private Sub lstFiles_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
On Error Resume Next
Dim CLMN_NAME As String, HDR_TITLE As String, E As Integer
 lstFiles.Sorted = True
 lstFiles.SortKey = ColumnHeader.Index - 1
 Select Case ColumnHeader.Index
 Case 1
        CLMN_NAME = "Name"
 Case 2
        CLMN_NAME = "Path"
 Case 3
        CLMN_NAME = "Type"
 Case 4
        CLMN_NAME = "Matches"
 End Select
 If lstFiles.SortOrder = lvwAscending Then
    lstFiles.SortOrder = lvwDescending
    ColumnHeader.Text = CLMN_NAME & " +"
 Else
    lstFiles.SortOrder = lvwAscending
    ColumnHeader.Text = CLMN_NAME & " -"
 End If
 For E = 1 To 4
   If Not E = ColumnHeader.Index Then
        HDR_TITLE = Mid(lstFiles.ColumnHeaders(E).Text, Len(lstFiles.ColumnHeaders(E).Text) - 1, Len(lstFiles.ColumnHeaders(E).Text))
        If HDR_TITLE = " +" Then lstFiles.ColumnHeaders(E).Text = Mid(lstFiles.ColumnHeaders(E).Text, 1, Len(lstFiles.ColumnHeaders(E).Text) - 1)
        If HDR_TITLE = " -" Then lstFiles.ColumnHeaders(E).Text = Mid(lstFiles.ColumnHeaders(E).Text, 1, Len(lstFiles.ColumnHeaders(E).Text) - 1)
   End If
 Next E
 lstFiles.Sorted = False
End Sub

'// SHOW POP-UP MENU FOR FILES
Private Sub lstFiles_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Me.PopupMenu mnuFileMenu
    End If
End Sub

'// RESIZE SOME CONTROLS
Private Sub Form_Resize()
On Error Resume Next
 If Not Me.WindowState = 1 Then
    Me.Width = 8085
    If Me.Height < 6810 Then Me.Height = 6810
    SSTab1.Height = Me.Height - SSTab1.Top - 800
    lstFiles.Height = SSTab1.Height - 3250
    RTF_LOG.Height = SSTab1.Height - 1500
    picHolder.Top = SSTab1.Height - 475
    pHolder1.Top = SSTab1.Height - 555
    pHolder3.Top = SSTab1.Height - 555
 End If
End Sub

'// SET FOCUS TO 1ST ITEM ON SELECTED TAB
Private Sub SSTab1_Click(PreviousTab As Integer)
On Error Resume Next
    If SSTab1.Tab = 0 Then coSearch.SetFocus
    If SSTab1.Tab = 1 Then chCase.SetFocus
    If SSTab1.Tab = 2 Then RTF_LOG.SetFocus
End Sub

'// MENU BUTTONS ON THE GENERAL TAB (PROCESS/REPLACE)
Private Sub tlbGeneral_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' PROCESSS
            If CHECK_INP = False Then Exit Sub
            Me.MousePointer = 11
            lstFiles.ListItems.Clear
            LST_COUNTER = 0
            frmMain.Caption = "0"
            Call Add_New_Mask
            Call Add_New_Replace
            Call Add_New_Search
            Call Add_New_Path
            Call Process_Listing
            tlbGeneral.Buttons(2).Enabled = True
            RTF_LOG.Text = ""
            Me.MousePointer = 0
            frmMain.Caption = "Search and Replace by Europeum.net - " & frmMain.Caption
        
        Case 2 ' REPLACE
            Dim intCHECKED As Integer, I As Integer
            intCHECKED = 0
            If CHECK_INP = False Then Exit Sub
            '// CHECK THAT WE HAVE SOMETHING TO PROCESS
            If Not lstFiles.ListItems.Count = 0 Then
                '// CHECK WHAT'S SELECTED
                For I = 1 To lstFiles.ListItems.Count
                    If lstFiles.ListItems(I).Checked = True Then intCHECKED = intCHECKED + 1
                Next I
                If intCHECKED = 0 Then
                    MsgBox "Please check files that you would like to search and replace.", vbInformation, "Checking ..."
                    Exit Sub
                End If
                Me.MousePointer = 11
                '// PROCEED
                If SHOW_VER = 1 Then
                    If MsgBox("Would you like to continue with the following settings:" & vbCrLf & vbCrLf & Create_Message, vbYesNo, "Verification ...") = vbYes Then
                        Call Replace_Matches
                        SSTab1.Tab = 2  '// SHOW REPORT
                        tlbGeneral.Buttons(2).Enabled = False
                    End If
                Else
                        Call Replace_Matches
                        SSTab1.Tab = 2  '// SHOW REPORT
                        tlbGeneral.Buttons(2).Enabled = False
                End If
            Else
                MsgBox "Please make sure to process the files first.", vbInformation, "Checking ..."
            End If '// 1st IF
            Me.MousePointer = 0
    End Select
End Sub

'// BUTTONS ON SETTINGS TAB
Private Sub tlbSettings_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 ' SAVE
            If frmMain.chVer.Value = 1 Then
             Call Save_Settings(True)
            Else
             Call Save_Settings(False)
            End If
        Case 2 ' DEFAULT
            If MsgBox("Set all settings to default ?", vbOKCancel, "Apply default settings ?") = vbOK Then
                chCase.Value = 0
                chFind.Value = 1
                chKeep.Value = 1
                optHTML.Value = True
                chSub.Value = 0
                chVer.Value = 1
                chWords.Value = 0
                chMulti.Value = 1
                chTouch.Value = 0
                chA.Value = 1
                chH.Value = 0
                chR.Value = 0
                txtPath.Text = "c:\"
                chSM.Value = 1
                pCOLOR.BackColor = 128
                pCOLORun.BackColor = &HE0E0E0
                chSplash.Value = 1
                If frmMain.chVer.Value = 1 Then
                 Call Save_Settings(True)
                Else
                 Call Save_Settings(False)
                End If
                
            End If
    End Select
End Sub

'// SAVE & PRINT REPORT
Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
Select Case Button.Index
    Case 1
        With CDC
            .DialogTitle = "Save report as ..."
             If AS_HTML = 1 Then
             .DefaultExt = "html"
             .Filter = "*.html"
             Else
             .DefaultExt = "txt"
             .Filter = "*.txt"
             End If
            .ShowSave
            If Not Trim(.FileName) = "" Then Call SaveFile(.FileName, RTF_LOG.Text)
        End With
    Case 2
        With CDC
            .DialogTitle = "Print Report Log ..."
            .CancelError = True
            .Flags = cdlPDReturnDC + cdlPDNoPageNums
            If RTF_LOG.SelLength = 0 Then
                .Flags = .Flags + cdlPDAllPages
            Else
                .Flags = .Flags + cdlPDSelection
            End If
            .ShowPrinter
            If ERR <> MSComDlg.cdlCancel Then
                RTF_LOG.SelPrint .hDC
            End If
        End With
End Select
End Sub






'//////  MENU ACTIONS \\\\\\
Private Sub mnuCheckUncheck_Click()
On Error GoTo ERR:
    If lstFiles.SelectedItem.Checked = True Then
        lstFiles.SelectedItem.Checked = False
    Else
        lstFiles.SelectedItem.Checked = True
    End If
Exit Sub
ERR:
    If ERR.Number = 91 Then Exit Sub
End Sub
Private Sub mnuCopy_Click()
Dim I As Integer, Copy_Processes As String
    If lstFiles.ListItems.Count = 0 Then
        MsgBox "Nothing to copy.", vbInformation, "Clipboard ..."
        Exit Sub
    End If
    For I = 1 To lstFiles.ListItems.Count
        Copy_Processes = Copy_Processes & lstFiles.ListItems(I).Text & Space(8) & lstFiles.ListItems(I).ListSubItems(1).Text & Space(8) & lstFiles.ListItems(I).ListSubItems(2).Text & Space(8) & lstFiles.ListItems(I).ListSubItems(3).Text & vbCrLf
    Next I
    Clipboard.SetText (Copy_Processes)
End Sub
Private Sub mnuPrinResults_Click()
On Error Resume Next
Dim Copy_Processes As String, I As Integer
        If lstFiles.ListItems.Count = 0 Then
            MsgBox "Nothing to print.", vbInformation, "Results ..."
            Exit Sub
        End If
        For I = 1 To lstFiles.ListItems.Count
            Copy_Processes = Copy_Processes & lstFiles.ListItems(I).Text & Space(8) & lstFiles.ListItems(I).ListSubItems(1).Text & Space(8) & lstFiles.ListItems(I).ListSubItems(2).Text & Space(8) & lstFiles.ListItems(I).ListSubItems(3).Text & vbCrLf
        Next I
        rtfHIDE.Text = Copy_Processes
        With CDC
            .DialogTitle = "Print Report Log ..."
            .CancelError = True
            .Flags = cdlPDReturnDC + cdlPDNoPageNums
            .Flags = .Flags + cdlPDSelection
            .ShowPrinter
            If ERR <> MSComDlg.cdlCancel Then
                rtfHIDE.SelPrint .hDC
            End If
        End With
End Sub
Private Sub mnuExit_Click()
    Unload Me
End Sub
Private Sub mnuExport_Click()
On Error Resume Next
Dim FL_N As String, CONTENT As String
    With CDC
        .DialogTitle = "Save export file as ..."
        .DefaultExt = "sar"
        .Filter = "*.sar"
        .ShowSave
        FL_N = .FileName
    End With
    If Not FL_N = "" Then
        CONTENT = CONTENT & Open_File(App.PATH & "\custom.ini", True) & "*******/////*******" & vbCrLf
        CONTENT = CONTENT & Open_File(App.PATH & "\srch.dat", True) & "*******/////*******" & vbCrLf
        CONTENT = CONTENT & Open_File(App.PATH & "\repl.dat", True) & "*******/////*******" & vbCrLf
        CONTENT = CONTENT & Open_File(App.PATH & "\path.dat", True) & "*******/////*******"
        CONTENT = CONTENT & Open_File(App.PATH & "\favorites.dat", True) & "*******/////*******" & vbCrLf
        CONTENT = CONTENT & Open_File(App.PATH & "\ext.dat", True)
        Call SaveFile(FL_N, CONTENT)
        MsgBox "Current environment has been exported to:" & vbCrLf & FL_N, vbInformation, "Export ..."
    End If
End Sub
Private Sub mnuImport_Click()
On Error Resume Next
Dim FL_N As String, CONTENT As String, EXP_UQ As Variant, I As Integer
    With CDC
        .DialogTitle = "Import ..."
        .DefaultExt = "sar"
        .Filter = "*.sar"
        .ShowOpen
        FL_N = .FileName
    End With
    If Not FL_N = "" Then
        If MsgBox("Are you sure you want to import following environment ?", vbYesNo, "Import ...") = vbYes Then
           EXP_UQ = Split(Open_File(FL_N, True), "*******/////*******")
           Call SaveFile(App.PATH & "\custom.ini", CStr(EXP_UQ(0)))     ' CUSTOM.INI
           Call SaveFile(App.PATH & "\srch.dat", CStr(EXP_UQ(1)))       ' SRCH.DAT
           Call SaveFile(App.PATH & "\repl.dat", CStr(EXP_UQ(2)))       ' REPL.DAT
           Call SaveFile(App.PATH & "\path.dat", CStr(EXP_UQ(3)))       ' PATH.DAT
           Call SaveFile(App.PATH & "\favorites.dat", CStr(EXP_UQ(4)))  ' FAVORITES.dat
           Call SaveFile(App.PATH & "\ext.dat", CStr(EXP_UQ(5)))        ' EXT.DAT
           'LOAD IMPORTED ENVIRONMENT
           Call Load_Settings(True)
           Call Load_Masks
           Call Load_Search
           Call Load_Replace
           Call Load_Path
           Call LOAD_FAVORITES_MENU
        End If
    End If
End Sub
Private Sub mnuOpenAssoc_Click()
On Error GoTo ERR:
    If Not lstFiles.ListItems.Count = 0 Then Shell ("start " & lstFiles.SelectedItem.ListSubItems(1).Text & lstFiles.SelectedItem.Text), vbHide
Exit Sub
ERR:
    MsgBox "Error: " & ERR.Number & vbCrLf & ERR.Description, , "Error ..."
End Sub
Private Sub mnuOrganizeFav_Click()
    frmORGFAV.Show 1, Me
End Sub
Private Sub mnuReport_Click()
    SSTab1.Tab = 2
End Sub
Private Sub mnuSettings_Click()
    SSTab1.Tab = 1
End Sub
Private Sub mnuAboutAbout_Click()
    frMAbout.Show 1, Me
End Sub
Private Sub mnuContents_Click()
On Error Resume Next
         Dim hwndHelp As Long
         hwndHelp = HtmlHelp(hWnd, App.PATH & "\hlp\sar.chm", HH_DISPLAY_TOPIC, 0)
End Sub
Private Sub mnuAddFavorite_Click()
    frmAddFavo.Show 1, Me
End Sub
Private Sub mnuViewFile_Click()
    On Error Resume Next
    FILE_TO_OPEN = lstFiles.SelectedItem.ListSubItems(1).Text & lstFiles.SelectedItem.Text
    frmEdit.Show 1, Me
End Sub
Private Sub mnusep6_Click(Index As Integer)
On Error Resume Next
    Call LOAD_NAME_FAVORITES(mnusep6(Index).Caption)
End Sub
Private Sub mnuViewRESHTML_Click()
On Error Resume Next
Dim I As Integer, Copy_Processes As String, CONTENT As String
    If lstFiles.ListItems.Count = 0 Then
        MsgBox "Nothing to view.", vbInformation, "Clipboard ..."
        Exit Sub
    End If
    CONTENT = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN'><html><head><title>Search And Replace Processed Results Log</title><style type='text/css'>body,p,td{font-family:Verdana, Arial, Helvetica, sans-serif;font-size:13px;color: Black}a:link    {color:navy;text-decoration:underline}a:visited {color:navy;text-decoration:underline}a:hover   {color:blue;text-decoration:normal}</style></head><body bgcolor='#ffffff'><table width='95%' align='center' cellpadding='2' cellspacing='0' border='1' bordercolor='WhiteSmoke'><tr bgcolor='whitesmoke'><td style='font size: 10px;'>ID</td><td>Name</td><td>Path</td><td>Type</td><td>Matches</td></tr>" & vbCrLf
    For I = 1 To lstFiles.ListItems.Count
        Copy_Processes = Copy_Processes & lstFiles.ListItems(I).Text & Space(8) & lstFiles.ListItems(I).ListSubItems(1).Text & Space(8) & lstFiles.ListItems(I).ListSubItems(2).Text & Space(8) & lstFiles.ListItems(I).ListSubItems(3).Text & vbCrLf
        CONTENT = CONTENT & "<tr><td><b>" & I & "</b></td>" & vbCrLf
        CONTENT = CONTENT & "<td>" & lstFiles.ListItems(I).Text & "</td>" & vbCrLf
        CONTENT = CONTENT & "<td>" & lstFiles.ListItems(I).ListSubItems(1).Text & "</td>" & vbCrLf
        CONTENT = CONTENT & "<td>" & lstFiles.ListItems(I).ListSubItems(2).Text & "</td>" & vbCrLf
        CONTENT = CONTENT & "<td>" & lstFiles.ListItems(I).ListSubItems(3).Text & "</td></tr>" & vbCrLf
    Next I
    CONTENT = CONTENT & "<tr bgcolor='WhiteSmoke'><td colspan='5'><font size='1'>Copyright &copy; 2002 <a href='http://www.Europeum.net' style='text-decoration:none;'>http://www.Europeum.net</a></font></td></tr></table></body></html>" & vbCrLf
    Call SaveFile(App.PATH & "\view_results~sar.html", CONTENT)
    Shell ("start explorer.exe " & App.PATH & "\view_results~sar.html"), vbMaximizedFocus
End Sub
