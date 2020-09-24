Attribute VB_Name = "MOD_INI"
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
' PS: This module has been coded by someone else (I have no idea     '
'     who, but it is available on PSC                                '
'                                                                    '
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'
'~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~'

Option Explicit
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Function mfncGetFromIni(strSectionHeader As String, strVariableName As String, strFilename As String) As String
    Dim strReturn As String
    strReturn = String(255, chR(0))
    mfncGetFromIni = Left$(strReturn, GetPrivateProfileString(strSectionHeader, ByVal strVariableName, "", strReturn, Len(strReturn), strFilename))
End Function
Function mfncParseString(strIn As String, intOffset As Integer, strDelimiter As String) As String
    If Len(strIn) = 0 Or intOffset = 0 Then
        mfncParseString = ""
        Exit Function
    End If
    Dim intStartPos As Integer
    ReDim intDelimPos(10) As Integer
    Dim intStrLen As Integer
    Dim intNoOfDelims As Integer
    Dim intCount As Integer
    Dim strQuotationMarks As String
    Dim intInsideQuotationMarks As Integer
    strQuotationMarks = chR(34) & chR(147) & chR(148)
    intInsideQuotationMarks = False
    For intCount = 1 To Len(strIn)
        If InStr(strQuotationMarks, Mid$(strIn, intCount, 1)) <> 0 Then
            intInsideQuotationMarks = (Not intInsideQuotationMarks)
        End If
        If (Not intInsideQuotationMarks) And (Mid$(strIn, intCount, 1) = strDelimiter) Then
        intNoOfDelims = intNoOfDelims + 1
        If (intNoOfDelims Mod 10) = 0 Then
            ReDim Preserve intDelimPos(intNoOfDelims + 10)
        End If
        intDelimPos(intNoOfDelims) = intCount
    End If
Next intCount
If intOffset > (intNoOfDelims + 1) Then
    mfncParseString = ""
    Exit Function
End If
If intOffset = 1 Then
    intStartPos = 1
End If
If intOffset = (intNoOfDelims + 1) Then
    If Right$(strIn, 1) = strDelimiter Then
        intStartPos = -1
        intStrLen = -1
        mfncParseString = ""
        Exit Function
    Else
        intStrLen = Len(strIn) - intDelimPos(intOffset - 1)
    End If
End If
If intStartPos = 0 Then
    intStartPos = intDelimPos(intOffset - 1) + 1
End If
If intStrLen = 0 Then
    intStrLen = intDelimPos(intOffset) - intStartPos
End If
mfncParseString = Mid$(strIn, intStartPos, intStrLen)
End Function
Function mfncWriteIni(strSectionHeader As String, strVariableName As String, strValue As String, strFilename As String) As Integer
    mfncWriteIni = WritePrivateProfileString(strSectionHeader, strVariableName, strValue, strFilename)
End Function


