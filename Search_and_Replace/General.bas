Attribute VB_Name = "General"
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

'// API Used for calling the chm help file
Declare Function HtmlHelp Lib "hhctrl.ocx" Alias "HtmlHelpA" (ByVal hwndCaller As Long, ByVal pszFile As String, ByVal uCommand As Long, ByVal dwData As Long) As Long
Public CASE_SENS As Integer, FIND_ALL As Integer, KEEP_LOG As Integer, AS_HTML As Integer, ROOT_PATH As String, SUB_DIR As Integer, SHOW_VER As Integer, WHOLE_W As Integer, BLOCK_SEARCH As Boolean, MULTI_LINE As Integer, TOUCH As Integer, TOUCH_CODE As Integer, SHOW_MATCHES As Integer
Public USE_TAB As Integer, LST_COUNTER As Integer, REPORT_SECTION As String, TIMER_C As Long, FOUND_CLR As Variant, NOTFOUND_CLR As Variant, SHOW_SPLASH As Integer
Public FILE_TO_OPEN As String
'// USED FOR THE FAVORITES
Public Type DataInfo
    strNAME As String * 50
    strSEARCH As String * 200
    strREPLACE As String * 200
    strMASK As String * 25
    strPATH As String * 250
    strSETTINGS As String * 18
    strTOUCH As String * 6
    strCOLOR As String * 60
    strROOT As String * 200
End Type

'// GET FILE ATTRIBS
Function GET_ATTR(ByRef PATH As String) As Integer
Dim fso As Variant, att As Variant
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set att = fso.GetFile(PATH)
    GET_ATTR = att.Attributes
    Set fso = Nothing
    Set att = Nothing
End Function

'// CREATE ATTRIB CODE FROM SETTINGS
Sub GET_TOUCH_CODE()
            ' Read-Only
            If frmMain.chR.Value = 1 And frmMain.chA.Value = 0 And frmMain.chH.Value = 0 Then TOUCH_CODE = 1
            ' Read-Only / Hidden
            If frmMain.chR.Value = 1 And frmMain.chA.Value = 0 And frmMain.chH.Value = 1 Then TOUCH_CODE = 3
            'Read-Only / Hidden / Archive
            If frmMain.chR.Value = 1 And frmMain.chA.Value = 1 And frmMain.chH.Value = 1 Then TOUCH_CODE = 35
            ' Read-Only / Archive
            If frmMain.chR.Value = 1 And frmMain.chA.Value = 1 And frmMain.chH.Value = 0 Then TOUCH_CODE = 33
            ' Hidden / Archive
            If frmMain.chR.Value = 0 And frmMain.chA.Value = 1 And frmMain.chH.Value = 1 Then TOUCH_CODE = 34
            ' Hidden
            If frmMain.chR.Value = 0 And frmMain.chA.Value = 0 And frmMain.chH.Value = 1 Then TOUCH_CODE = 2
            ' Archive
            If frmMain.chR.Value = 0 And frmMain.chA.Value = 1 And frmMain.chH.Value = 0 Then TOUCH_CODE = 32
            ' None
            If frmMain.chR.Value = 0 And frmMain.chA.Value = 0 And frmMain.chH.Value = 0 Then TOUCH_CODE = 0
End Sub

'// APPLY SETTINGS FROM DATA
Sub APPLY_TOUCH_CODE(ByRef CODE As Integer)
    Select Case CODE
        Case 1             ' Read-Only
            frmMain.chR.Value = 1
            frmMain.chA.Value = 0
            frmMain.chH.Value = 0
        Case 3             ' Read-Only / Hidden
            frmMain.chR.Value = 1
            frmMain.chA.Value = 0
            frmMain.chH.Value = 1
        Case 35             'Read-Only / Hidden / Archive
            frmMain.chR.Value = 1
            frmMain.chA.Value = 1
            frmMain.chH.Value = 1
        Case 33            ' Read-Only / Archive
            frmMain.chR.Value = 1
            frmMain.chA.Value = 1
            frmMain.chH.Value = 0
        Case 34            ' Hidden / Archive
            frmMain.chR.Value = 0
            frmMain.chA.Value = 1
            frmMain.chH.Value = 1
        Case 2             ' Hidden
            frmMain.chR.Value = 0
            frmMain.chA.Value = 0
            frmMain.chH.Value = 1
        Case 32            ' Archive
            frmMain.chR.Value = 0
            frmMain.chA.Value = 1
            frmMain.chH.Value = 0
        Case 0             ' None
            frmMain.chR.Value = 0
            frmMain.chA.Value = 0
            frmMain.chH.Value = 0
    End Select
End Sub

'// LOADS SETTINGS TO VARIABLES AND IF APPLIED TO THE SETTINGS TAB
Sub Load_Settings(ByRef APPLY As Boolean)
    CASE_SENS = CInt(mfncGetFromIni("GENERAL", "CASE", App.PATH & "\custom.ini"))
    FIND_ALL = CInt(mfncGetFromIni("GENERAL", "FIND_ALL", App.PATH & "\custom.ini"))
    KEEP_LOG = CInt(mfncGetFromIni("GENERAL", "KEEP_LOG", App.PATH & "\custom.ini"))
    AS_HTML = CInt(mfncGetFromIni("GENERAL", "AS_HTML", App.PATH & "\custom.ini"))
    ROOT_PATH = mfncGetFromIni("GENERAL", "ROOT_PATH", App.PATH & "\custom.ini")
    SUB_DIR = mfncGetFromIni("GENERAL", "SUB_DIR", App.PATH & "\custom.ini")
    SHOW_VER = mfncGetFromIni("GENERAL", "SHOW_VER", App.PATH & "\custom.ini")
    WHOLE_W = mfncGetFromIni("GENERAL", "WHOLE_W", App.PATH & "\custom.ini")
    MULTI_LINE = mfncGetFromIni("GENERAL", "MULTI_LINE", App.PATH & "\custom.ini")
    TOUCH_CODE = CInt(mfncGetFromIni("GENERAL", "TOUCH_CODE", App.PATH & "\custom.ini"))
    TOUCH = mfncGetFromIni("GENERAL", "TOUCH", App.PATH & "\custom.ini")
    SHOW_MATCHES = mfncGetFromIni("GENERAL", "SHOW_MATCHES", App.PATH & "\custom.ini")
    FOUND_CLR = mfncGetFromIni("GENERAL", "FOUND_CLR", App.PATH & "\custom.ini")
    NOTFOUND_CLR = mfncGetFromIni("GENERAL", "NOTFOUND_CLR", App.PATH & "\custom.ini")
    SHOW_SPLASH = mfncGetFromIni("GENERAL", "SHOW_SPLASH", App.PATH & "\custom.ini")
    If APPLY = True Then Call Apply_Settings
End Sub

'// APPLIES SETTINGS TO THE SETTINGS TAB - CALLED FROM 'Load_Settings'
Sub Apply_Settings()
    If CASE_SENS = 1 Then frmMain.chCase.Value = 1
    If FIND_ALL = 1 Then frmMain.chFind.Value = 1
    If KEEP_LOG = 1 Then frmMain.chKeep.Value = 1
    If SUB_DIR = 1 Then frmMain.chSub.Value = 1
    If SHOW_VER = 1 Then frmMain.chVer.Value = 1
    If WHOLE_W = 1 Then frmMain.chWords.Value = 1
    If MULTI_LINE = 1 Then frmMain.chMulti.Value = 1
    If TOUCH = 1 Then frmMain.chTouch.Value = 1
    If SHOW_MATCHES = 1 Then frmMain.chSM.Value = 1
    If SHOW_SPLASH = 1 Then frmMain.chSplash.Value = 1
    frmMain.pCOLOR.BackColor = Val(FOUND_CLR)
    frmMain.pCOLORun.BackColor = Val(NOTFOUND_CLR)
    frmMain.txtPath.Text = ROOT_PATH
    If AS_HTML = 1 Then
        frmMain.optHTML.Value = True
        frmMain.optText.Value = False
    Else
        frmMain.optHTML.Value = False
        frmMain.optText.Value = True
    End If
    Call APPLY_TOUCH_CODE(Int(TOUCH_CODE))
    Call frmMain.chTouch_Click
End Sub

'// SAVE SETTING TO INI AND VARIABLES - IF SHOW_MSG SET TO TRUE THEN SHOW VERIFICATION
Sub Save_Settings(ByRef SHOW_MSG As Boolean)
    CASE_SENS = frmMain.chCase.Value
    FIND_ALL = frmMain.chFind.Value
    KEEP_LOG = frmMain.chKeep.Value
    ROOT_PATH = frmMain.txtPath.Text
    SUB_DIR = frmMain.chSub.Value
    SHOW_VER = frmMain.chVer.Value
    WHOLE_W = frmMain.chWords.Value
    MULTI_LINE = frmMain.chMulti.Value
    TOUCH = frmMain.chTouch.Value
    SHOW_MATCHES = frmMain.chSM.Value
    FOUND_CLR = frmMain.pCOLOR.BackColor
    NOTFOUND_CLR = frmMain.pCOLORun.BackColor
    SHOW_SPLASH = frmMain.chSplash.Value
    Call GET_TOUCH_CODE
    If frmMain.optHTML.Value = True Then AS_HTML = 1
    If frmMain.optHTML.Value = False Then AS_HTML = 0
    
    Call mfncWriteIni("GENERAL", "TOUCH", CStr(TOUCH), App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "TOUCH_CODE", CStr(TOUCH_CODE), App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "CASE", CStr(CASE_SENS), App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "FIND_ALL", CStr(FIND_ALL), App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "KEEP_LOG", CStr(KEEP_LOG), App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "AS_HTML", CStr(AS_HTML), App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "ROOT_PATH", ROOT_PATH, App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "SUB_DIR", CStr(SUB_DIR), App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "SHOW_VER", CStr(SHOW_VER), App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "WHOLE_W", CStr(WHOLE_W), App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "MULTI_LINE", CStr(MULTI_LINE), App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "SHOW_MATCHES", CStr(SHOW_MATCHES), App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "FOUND_CLR", CStr(FOUND_CLR), App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "NOTFOUND_CLR", CStr(NOTFOUND_CLR), App.PATH & "\custom.ini")
    Call mfncWriteIni("GENERAL", "SHOW_SPLASH", CStr(SHOW_SPLASH), App.PATH & "\custom.ini")
    If SHOW_MSG = True Then MsgBox "Settings have been sucesfully saved.", vbInformation, "Saved ..."
End Sub

'// ADDS NEW ITEM TO THE FILE IF IT NOT ALREADY EXISTS
Sub Add_New_Mask()
Dim FL_CONTENT As String
    FL_CONTENT = Open_File(App.PATH & "\ext.dat", False)
    If Not InStr(1, LCase(FL_CONTENT), LCase(frmMain.coMask.Text & "|")) >= 1 Then
        Call SaveFile(App.PATH & "\ext.dat", frmMain.coMask.Text & "|" & vbCrLf & FL_CONTENT)
    End If
End Sub
Sub Add_New_Search()
Dim FL_CONTENT As String
    FL_CONTENT = Open_File(App.PATH & "\srch.dat", False)
    If Not InStr(1, LCase(FL_CONTENT), LCase(frmMain.coSearch.Text & "delim|~|~|~|delim")) >= 1 Then
        Call SaveFile(App.PATH & "\srch.dat", frmMain.coSearch.Text & "delim|~|~|~|delim" & vbCrLf & FL_CONTENT)
    End If
End Sub
Sub Add_New_Replace()
Dim FL_CONTENT As String
    FL_CONTENT = Open_File(App.PATH & "\repl.dat", False)
    If Not InStr(1, LCase(FL_CONTENT), LCase(frmMain.coReplace.Text & "delim|~|~|~|delim")) >= 1 Then
        Call SaveFile(App.PATH & "\repl.dat", frmMain.coReplace.Text & "delim|~|~|~|delim" & vbCrLf & FL_CONTENT)
    End If
End Sub
Sub Add_New_Path()
Dim FL_CONTENT As String
    FL_CONTENT = Open_File(App.PATH & "\path.dat", False)
    If Not InStr(1, LCase(FL_CONTENT), LCase(frmMain.coPath.Text & "||")) >= 1 Then
        Call SaveFile(App.PATH & "\path.dat", frmMain.coPath.Text & "||" & FL_CONTENT)
    End If
End Sub

'// LOADS THE PREVIOUSLY SELECTED ITEMS TO THE COMBO BOX
Sub Load_Masks()
On Error Resume Next
Dim I As Integer, FILE_EXT As String, UNQ_EXT_LN As Variant
    frmMain.coMask.Clear
    FILE_EXT = Open_File(App.PATH & "\ext.dat", False)
    UNQ_EXT_LN = Split(FILE_EXT, "|")
    For I = 0 To UBound(UNQ_EXT_LN)
        If Not Trim(UNQ_EXT_LN(I)) = "" Then frmMain.coMask.AddItem UNQ_EXT_LN(I)
    Next I
End Sub
Sub Load_Search()
Dim I As Integer, FILE_EXT As String, UNQ_EXT_LN As Variant
    frmMain.coSearch.Clear
    FILE_EXT = Open_File(App.PATH & "\srch.dat", False)
    UNQ_EXT_LN = Split(FILE_EXT, "delim|~|~|~|delim")
    For I = 0 To UBound(UNQ_EXT_LN)
        If Not Trim(UNQ_EXT_LN(I)) = "" Then frmMain.coSearch.AddItem UNQ_EXT_LN(I)
    Next I
End Sub
Sub Load_Replace()
Dim I As Integer, FILE_EXT As String, UNQ_EXT_LN As Variant
    frmMain.coReplace.Clear
    FILE_EXT = Open_File(App.PATH & "\repl.dat", False)
    UNQ_EXT_LN = Split(FILE_EXT, "delim|~|~|~|delim")
    For I = 0 To UBound(UNQ_EXT_LN)
        If Not Trim(UNQ_EXT_LN(I)) = "" Then frmMain.coReplace.AddItem UNQ_EXT_LN(I)
    Next I
End Sub
Sub Load_Path()
Dim I As Integer, FILE_EXT As String, UNQ_EXT_LN As Variant
    frmMain.coPath.Clear
    FILE_EXT = Open_File(App.PATH & "\path.dat", False)
    UNQ_EXT_LN = Split(FILE_EXT, "||")
    For I = 0 To UBound(UNQ_EXT_LN)
        If Not Trim(UNQ_EXT_LN(I)) = "" Then frmMain.coPath.AddItem UNQ_EXT_LN(I)
    Next I
End Sub

'// CHECKS IF THE USER HAS ALL FIELDS ENTERD CORRECTLY
Function CHECK_INP() As Boolean
Dim ERR As String
    CHECK_INP = True
    If Trim(frmMain.coSearch.Text) = "" Then
        CHECK_INP = False
        ERR = ERR & "Please enter search string." & vbCrLf
    End If
    If Trim(frmMain.coReplace.Text) = "" Then
        CHECK_INP = False
        ERR = ERR & "Please enter replace string." & vbCrLf
    End If
    If Trim(frmMain.coPath.Text) = "" Then
        CHECK_INP = False
        ERR = ERR & "Please enter valid path." & vbCrLf
    End If
    If Trim(frmMain.coMask.Text) = "" Then
        CHECK_INP = False
        ERR = ERR & "Please enter desired file mask." & vbCrLf
    End If
    If CHECK_INP = False Then MsgBox ERR, vbInformation, "Missing ...."
End Function
