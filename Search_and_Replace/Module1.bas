Attribute VB_Name = "Other"
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

'// ADD LISTING TO THE SHOWN FILES
Sub ADD_LISTING(ByRef LST As ListView, NAME As String, ITEM1 As String, ITEM2 As String, ITEM3 As String)
LST_COUNTER = LST_COUNTER + 1
    LST.ListItems.Add LST_COUNTER, , NAME
    LST.ListItems(LST_COUNTER).SubItems(1) = ITEM1
    LST.ListItems(LST_COUNTER).SubItems(2) = ITEM2
    
    If SHOW_MATCHES = 1 Then
        LST.ListItems(LST_COUNTER).SubItems(3) = ITEM3
    Else
        LST.ListItems(LST_COUNTER).SubItems(3) = "N/A"
    End If
    
    If Not ITEM3 = "0" Or ITEM3 = "" Then
        LST.ListItems(LST_COUNTER).ForeColor = Val(FOUND_CLR)
        LST.ListItems(LST_COUNTER).ListSubItems(1).ForeColor = Val(FOUND_CLR)
        LST.ListItems(LST_COUNTER).ListSubItems(2).ForeColor = Val(FOUND_CLR)
        LST.ListItems(LST_COUNTER).ListSubItems(3).ForeColor = Val(FOUND_CLR)
    Else
        LST.ListItems(LST_COUNTER).ForeColor = Val(NOTFOUND_CLR)
        LST.ListItems(LST_COUNTER).ListSubItems(1).ForeColor = Val(NOTFOUND_CLR)
        LST.ListItems(LST_COUNTER).ListSubItems(2).ForeColor = Val(NOTFOUND_CLR)
        LST.ListItems(LST_COUNTER).ListSubItems(3).ForeColor = Val(NOTFOUND_CLR)
    End If
    
End Sub

'// DECIDE IF WE SHOULD INCLUDE SUB_FOLDERS IN THE PROCESSING
Sub Process_Listing()
    If SUB_DIR = 1 Then '// YES INCLUDE SUB-FOLDERS
        Call List_Simple(MODIFY_PATH(frmMain.coPath.Text))
        Call DirMap(MODIFY_PATH(frmMain.coPath.Text))
    Else                '// NO SUB-FOLDERS
        Call List_Simple(MODIFY_PATH(frmMain.coPath.Text))
    End If
End Sub

'// CHECK THAT THE PATH HAS '\' ON THE END AND THAT IT'S A VALID PATH
Function MODIFY_PATH(PATH As String) As String
    If Not Mid(PATH, Len(PATH), 1) = "\" Then PATH = PATH & "\"
    MODIFY_PATH = PATH
End Function

'// LIST ALL FILES MATCHING FILE MASK AND SHOW IT'S DETAILS
Sub List_Simple(ByRef PATH As String)
On Error Resume Next
Dim I As Integer
    frmMain.Dir1.PATH = PATH
    For I = 0 To frmMain.File1.ListCount
        If Not Trim(frmMain.File1.List(I)) = "" Then
            If Check_Ext(File_Extention(PATH & "\" & frmMain.File1.List(I))) = True Then
                If CHECK_ATTRIB(PATH & "\" & frmMain.File1.List(I)) = True Then
                    Call ADD_LISTING(frmMain.lstFiles, frmMain.File1.List(I), PATH, LCase(File_Extention(PATH & "\" & frmMain.File1.List(I))), COUNT_MATCHES(PATH & "\" & frmMain.File1.List(I)))
                End If
            End If
        End If
    Next I
End Sub

'// CHECK FOR FILE ATTRIBUTES
Function CHECK_ATTRIB(ByRef PATH As String) As Boolean
    CHECK_ATTRIB = False
    If TOUCH = 1 Then '// APPLY
        If GET_ATTR(PATH) = Int(TOUCH_CODE) Then '//SAME SETTINGS AS FILE
            CHECK_ATTRIB = True
        Else
            CHECK_ATTRIB = False
        End If
    Else
        CHECK_ATTRIB = True
    End If
End Function

'// LIST ONLY FILES WITH VALID EXTENTION
Function Check_Ext(ByRef EXT As String) As Boolean
Dim EXTS As String, EXTS_UQ As Variant, I As Integer
    Check_Ext = False
    EXTS = Replace(frmMain.coMask.Text, " ", "")

    If Trim(EXTS) = "" And Not Mid(EXTS, 1, 2) = "*." Then
        Check_Ext = False
        Exit Function
    End If
    
    If Not InStr(1, EXTS, ",") >= 1 Then EXTS = EXTS & ","
    EXTS_UQ = Split(EXTS, ",")
    For I = 0 To UBound(EXTS_UQ)
        If "*." & LCase(EXT) = LCase(EXTS_UQ(I)) Or LCase(EXTS_UQ(I)) = "*.*" Then Check_Ext = True
    Next I
End Function


'// COUNT HOW MANY MATCHES THERE ARE IN THE FILE !!AND!! FOLLOW USER SETTING
Function COUNT_MATCHES(ByRef PATH As String) As Integer
Dim re As New RegExp, ma As Match
COUNT_MATCHES = 0
If SHOW_MATCHES = 1 Then
    If Not WHOLE_W = 1 Then             '// ANY MATCH
        re.Pattern = frmMain.coSearch
    Else
        re.Pattern = "\b" & frmMain.coSearch & "\b"
    End If
    If CASE_SENS = 1 Then
        re.IgnoreCase = False                  ' case sensitive search
    Else
        re.IgnoreCase = True
    End If
    If FIND_ALL = 1 Then
        re.Global = True                       ' find all the occurrences
    Else
        re.Global = False
    End If
    ' THIS FOR NEXT STATEMENT IS WHAT SLOWS DOWN THE FILE LISINTG,
    ' IT HAS TO COUNT EACH MATCH AND ADD IT ONTO THE COUNTER
    For Each ma In re.Execute(Open_File(PATH, True))
        COUNT_MATCHES = COUNT_MATCHES + 1
    Next
End If
frmMain.Caption = Int(frmMain.Caption) + 1
End Function

'// LIST ALL FILES MATCHING FILE MASK AND SHOW IT'S DETAILS
Sub List_Extended_Add(ByRef PATH As String)
On Error Resume Next
Dim I As Integer
    frmMain.Dir1.PATH = PATH
    For I = 0 To frmMain.File1.ListCount
        If Not Trim(frmMain.File1.List(I)) = "" Then
            If Check_Ext(File_Extention(PATH & "\" & frmMain.File1.List(I))) = True Then
                If CHECK_ATTRIB(PATH & "\" & frmMain.File1.List(I)) = True Then
                    Call ADD_LISTING(frmMain.lstFiles, frmMain.File1.List(I), PATH, LCase(File_Extention(PATH & "\" & frmMain.File1.List(I))), COUNT_MATCHES(PATH & "\" & frmMain.File1.List(I)))
                End If
            End If
        End If
    Next I
End Sub

'// THANKS TO Coolwick (PSC) FOR THIS DIR SCANER
Sub DirMap(ByVal PATH As String)
On Error Resume Next
    Dim I, j, X As Integer, PASS_DIR As String
    Dim Fname(), CurrentFolder, Temp As String
    Temp = PATH
    If DIR(Temp, vbDirectory) = "" Then Exit Sub
    CurrentFolder = DIR(Temp, vbDirectory)
    Do While CurrentFolder <> ""
        If GetAttr(Temp & CurrentFolder) = vbDirectory Then
            If CurrentFolder <> "." And CurrentFolder <> ".." Then
                I = I + 1
            End If
        End If
        CurrentFolder = DIR
    Loop
    ReDim Fname(I)
    CurrentFolder = DIR(Temp, vbDirectory)
    Do While CurrentFolder <> ""
        If GetAttr(Temp & CurrentFolder) = vbDirectory Then
            If CurrentFolder <> "." And CurrentFolder <> ".." Then
                j = j + 1
                Fname(j) = CurrentFolder
                PASS_DIR = Temp & Fname(j)
                If Not Mid(PASS_DIR, Len(PASS_DIR), 1) = "\" Then PASS_DIR = PASS_DIR & "\"
                Call List_Extended_Add(PASS_DIR)
            End If
        End If
        CurrentFolder = DIR
    Loop
    For X = 1 To I
        Call DirMap(Temp & Fname(X) & "\")
    Next
End Sub

