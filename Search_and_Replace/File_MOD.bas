Attribute VB_Name = "File_MOD"
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

'// FILE API
Declare Function ShellExecuteEx& _
    Lib "Shell32.dll" Alias "ShellExecuteExA" (ByRef lpExecInfo As SHELLEXECUTEINFO)
    Const SEE_MASK_INVOKEIDLIST& = &HC
Type SHELLEXECUTEINFO
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

'// GET FILE DATE
Function Get_File_Date(ByRef FILE_PATH As String) As String
On Error Resume Next
    Get_File_Date = FileDateTime(FILE_PATH)
End Function

'// GET FILE SIZE IN (BYTES)
Function Get_File_Size(ByRef FILE_PATH As String) As String
On Error Resume Next
    Get_File_Size = CCur(FileLen(FILE_PATH))
End Function

'// SAVE FILE
Sub SaveFile(ByVal PATH As String, CONTENT As String)
On Error GoTo ERR:
Dim intFile As Integer
     intFile = FreeFile
    Open PATH For Output As #intFile
        Print #intFile, CONTENT
    Close #intFile
    Exit Sub
ERR:
If ERR.Number = 75 Then MsgBox "Error accesing file:" & vbCrLf & PATH & vbCrLf & "Please make sure that the file is not with a read only attribute.", vbCritical, "Error ..."
End Sub

'// APPEND TO FILE
Sub AppendFile(ByVal PATH As String, CONTENT As String)
On Error Resume Next
Dim intFile As Integer
        intFile = FreeFile
        Open PATH For Append As #intFile
            Print #intFile, CONTENT
        Close #intFile
End Sub

'// OPEN FILE AS FILE CONTENT
Function Open_File(ByRef PATH As String, NEW_LN As Boolean) As String
Dim intFileNum As Integer, strTextLine As String
    intFileNum = FreeFile
     Open PATH For Input As #intFileNum
      Do While Not EOF(intFileNum)
       Line Input #intFileNum, strTextLine
       If NEW_LN = True Then
         Open_File = Open_File & strTextLine & vbNewLine
       Else
         Open_File = Open_File & strTextLine
       End If
      Loop
    Close #intFileNum
End Function

'// GET FILE TYPE
Function File_Extention(ByRef PATH As String) As String
On Error Resume Next
    File_Extention = Mid(PATH, InStrRev(PATH, ".") + 1)
End Function

