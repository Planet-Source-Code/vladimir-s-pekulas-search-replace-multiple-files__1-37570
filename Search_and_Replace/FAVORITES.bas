Attribute VB_Name = "FAVORITES"
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

'// ADD FAVORITE TO FAVORITES.DAT FILE
Sub ADD_FOVORITE(ByRef NAME As String, SRCH As String, REPL As String, MASK As String, PATH As String, SETTINGS As String, TOUCH As String, COLOR As String, ROOT As String)
On Error Resume Next
Dim udtAddData As DataInfo, lngNextID As Long, intFile As Integer, NumRecords As Integer
Dim lngRecLength As Long, lngID As Integer
    intFile = FreeFile
    lngRecLength = LenB(udtAddData)
    Open App.PATH & "\favorites.dat" For Random As #intFile Len = lngRecLength
    If LOF(intFile) Mod lngRecLength = 0 Then
        NumRecords = (LOF(intFile) \ lngRecLength)
    Else
        NumRecords = (LOF(intFile) \ lngRecLength) + 1
    End If
     lngID = 0
      Do
       lngID = lngID + 1
         If lngID > NumRecords Then GoTo Done:
          Get #intFile, lngID, udtAddData
            If Trim(udtAddData.strNAME) = Trim(frmAddFavo.txtNAME.Text) Then
                MsgBox "The favorite name already exists.", vbInformation, "Save error"
                Close #intFile
                Exit Sub
            End If
      Loop

Done:
    lngID = NumRecords
    lngNextID = lngID + 1
    udtAddData.strNAME = NAME
    udtAddData.strSEARCH = SRCH
    udtAddData.strREPLACE = REPL
    udtAddData.strMASK = MASK
    udtAddData.strPATH = PATH
    udtAddData.strSETTINGS = SETTINGS
    udtAddData.strCOLOR = COLOR
    udtAddData.strROOT = ROOT
    Put #intFile, lngNextID, udtAddData
 Close #intFile
End Sub

'// LOAD FAVORITES TO LISTVIEW
Sub LOAD_FAVORITES()
On Error Resume Next
Dim udtLoadData As DataInfo, lngNextID As Long, intFile As Integer, lngRecLength As Long, NumRecords As Integer, LST_COUNTER_FAV As Integer
Dim lngID As Integer
    intFile = FreeFile
    lngRecLength = LenB(udtLoadData)
    Open App.PATH & "\favorites.dat" For Random As #intFile Len = lngRecLength
    If LOF(intFile) Mod lngRecLength = 0 Then
        NumRecords = (LOF(intFile) \ lngRecLength)
    Else
        NumRecords = (LOF(intFile) \ lngRecLength) + 1
    End If
    lngID = 0
  Do
   lngID = lngID + 1
     If lngID > NumRecords Then GoTo Done:
      Get #intFile, lngID, udtLoadData
        If Not Trim(udtLoadData.strNAME) = "" Then
            LST_COUNTER_FAV = LST_COUNTER_FAV + 1
            frmORGFAV.lstLIST.ListItems.Add LST_COUNTER_FAV, , Trim(udtLoadData.strNAME)
            frmORGFAV.lstLIST.ListItems(LST_COUNTER_FAV).SubItems(1) = Trim(udtLoadData.strSEARCH)
            frmORGFAV.lstLIST.ListItems(LST_COUNTER_FAV).SubItems(2) = Trim(udtLoadData.strREPLACE)
            frmORGFAV.lstLIST.ListItems(LST_COUNTER_FAV).SubItems(3) = Trim(udtLoadData.strMASK)
            frmORGFAV.lstLIST.ListItems(LST_COUNTER_FAV).SubItems(4) = Trim(udtLoadData.strPATH)
            frmORGFAV.lstLIST.ListItems(LST_COUNTER_FAV).SubItems(5) = lngID
        End If
  Loop
Close #intFile
Exit Sub
Done:
 Close #intFile
End Sub

'// LOAD FAVORITES TO MENU
Sub LOAD_FAVORITES_MENU()
On Error Resume Next
Dim udtLoadData As DataInfo, lngNextID As Long, intFile As Integer, lngRecLength As Long, NumRecords As Integer, LST_COUNTER_FAV As Integer, I As Integer
Dim lngID As Integer
    intFile = FreeFile
    lngRecLength = LenB(udtLoadData)
    Open App.PATH & "\favorites.dat" For Random As #intFile Len = lngRecLength
    If LOF(intFile) Mod lngRecLength = 0 Then
        NumRecords = (LOF(intFile) \ lngRecLength)
    Else
        NumRecords = (LOF(intFile) \ lngRecLength) + 1
    End If
    lngID = 0
    I = 0
  Do
   lngID = lngID + 1
     If lngID > NumRecords Then GoTo Done:
      Get #intFile, lngID, udtLoadData
       If Not Trim(udtLoadData.strNAME) = "" Then
            I = I + 1
            Load frmMain.mnusep6(I)
            frmMain.mnusep6(I).Caption = Trim(udtLoadData.strNAME)
       End If
  Loop
Close #intFile
Exit Sub
Done:
 Close #intFile
End Sub

'// EMPTY THE FAVORITES MENU
Sub EMPTY_FAV_MENU()
Dim I As Integer
    For I = 1 To frmMain.mnusep6.Count - 1
        Unload frmMain.mnusep6(I)
    Next I
End Sub

'// APPLY FAVORITE DETAILS TO THE APPLICATION
Sub LOAD_ID_FAVORITES(ByRef ID As Integer)
On Error Resume Next
Dim udtLoadData As DataInfo, intFile As Integer, lngRecLength As Long, NumRecords As Integer
Dim strTOUCH As String, strSETTINGS As String, strCOLORS As String, CO_ARRAY_BC As Variant, CO_ARRAY_T As Variant
    intFile = FreeFile
    lngRecLength = LenB(udtLoadData)
    Open App.PATH & "\favorites.dat" For Random As #intFile Len = lngRecLength
    If LOF(intFile) Mod lngRecLength = 0 Then
        NumRecords = (LOF(intFile) \ lngRecLength)
    Else
        NumRecords = (LOF(intFile) \ lngRecLength) + 1
    End If
    Get #intFile, ID, udtLoadData
        frmMain.coSearch.Text = Trim(udtLoadData.strSEARCH)
        frmMain.coReplace.Text = Trim(udtLoadData.strREPLACE)
        frmMain.coMask.Text = Trim(udtLoadData.strMASK)
        frmMain.coPath.Text = Trim(udtLoadData.strPATH)
        strCOLORS = Trim(udtLoadData.strCOLOR)
        strSETTINGS = Trim(udtLoadData.strSETTINGS)
    Close #intFile
    '// COLORS
    CO_ARRAY_BC = Split(strCOLORS, ",")
    frmMain.pCOLOR.BackColor = CO_ARRAY_BC(0)
    frmMain.pCOLORun.BackColor = CO_ARRAY_BC(1)
    '// TOUCH
    CO_ARRAY_T = Split(strSETTINGS, ",")
    '// SETTINGS
    frmMain.chCase.Value = CO_ARRAY_T(0)
    frmMain.chFind.Value = CO_ARRAY_T(1)
    frmMain.chKeep.Value = CO_ARRAY_T(2)
    frmMain.chSub.Value = CO_ARRAY_T(3)
    frmMain.chVer.Value = CO_ARRAY_T(4)
    frmMain.chWords.Value = CO_ARRAY_T(5)
    frmMain.chMulti.Value = CO_ARRAY_T(6)
    frmMain.chTouch.Value = CO_ARRAY_T(7)
    frmMain.chSM.Value = CO_ARRAY_T(8)
    '// SAVE SETTINGS
    Call Save_Settings(False)
End Sub

'// APPLY FAVORITE DETAILS TO THE APPLICATION
Sub LOAD_NAME_FAVORITES(ByRef NAME As String)
On Error Resume Next
Dim udtLoadData As DataInfo, intFile As Integer, lngRecLength As Long, NumRecords As Integer
Dim strTOUCH As String, strSETTINGS As String, strCOLORS As String, CO_ARRAY_BC As Variant, CO_ARRAY_T As Variant, I As Integer
    intFile = FreeFile
    lngRecLength = LenB(udtLoadData)
    Open App.PATH & "\favorites.dat" For Random As #intFile Len = lngRecLength
    If LOF(intFile) Mod lngRecLength = 0 Then
        NumRecords = (LOF(intFile) \ lngRecLength)
    Else
        NumRecords = (LOF(intFile) \ lngRecLength) + 1
    End If
    
    For I = 1 To NumRecords
        Get #intFile, I, udtLoadData
        If UCase(Trim(NAME)) = UCase(Trim(udtLoadData.strNAME)) Then
            frmMain.coSearch.Text = Trim(udtLoadData.strSEARCH)
            frmMain.coReplace.Text = Trim(udtLoadData.strREPLACE)
            frmMain.coMask.Text = Trim(udtLoadData.strMASK)
            frmMain.coPath.Text = Trim(udtLoadData.strPATH)
            strCOLORS = Trim(udtLoadData.strCOLOR)
            strSETTINGS = Trim(udtLoadData.strSETTINGS)
            '// COLORS
            CO_ARRAY_BC = Split(strCOLORS, ",")
            frmMain.pCOLOR.BackColor = CO_ARRAY_BC(0)
            frmMain.pCOLORun.BackColor = CO_ARRAY_BC(1)
            '// TOUCH
            CO_ARRAY_T = Split(strSETTINGS, ",")
            '// SETTINGS
            frmMain.chCase.Value = CO_ARRAY_T(0)
            frmMain.chFind.Value = CO_ARRAY_T(1)
            frmMain.chKeep.Value = CO_ARRAY_T(2)
            frmMain.chSub.Value = CO_ARRAY_T(3)
            frmMain.chVer.Value = CO_ARRAY_T(4)
            frmMain.chWords.Value = CO_ARRAY_T(5)
            frmMain.chMulti.Value = CO_ARRAY_T(6)
            frmMain.chTouch.Value = CO_ARRAY_T(7)
            frmMain.chSM.Value = CO_ARRAY_T(8)
            '// SAVE SETTINGS
            Call Save_Settings(False)
        End If
    Next I
    Close #intFile
End Sub

'// DELETE FAVORITE FROM FAVORITES.DAT
'// THIS ACCTUALLY SETS THE VALUES TO "" RATHER THEN DELETING THE RECORD
'// IT SELF, SINCE UDT DOESN'T OFFER ANY 'BUILD-IN' WAY OF DOING SO, YOU'D
'// HAVE TO REBUILD THE FILE IT SELF WHILE LEAVING SELECTED RECORD OUT.
'// FOR OUR PURPOSE THIS SHOULD BE ENOUGHT :)
Sub DELETE_FOVORITE(ByRef ID As Integer)
On Error Resume Next
Dim udtAddData As DataInfo, lngNextID As Long, intFile As Integer, NumRecords As Integer, lngRecLength As Long, lngID As Integer
    intFile = FreeFile
    lngRecLength = LenB(udtAddData)
    Open App.PATH & "\favorites.dat" For Random As #intFile Len = lngRecLength
    If LOF(intFile) Mod lngRecLength = 0 Then
        NumRecords = (LOF(intFile) \ lngRecLength)
    Else
        NumRecords = (LOF(intFile) \ lngRecLength) + 1
    End If
     lngID = ID
     If lngID > NumRecords Then GoTo Done:
        lngID = ID
        udtAddData.strNAME = ""
        udtAddData.strSEARCH = ""
        udtAddData.strREPLACE = ""
        udtAddData.strMASK = ""
        udtAddData.strPATH = ""
        udtAddData.strSETTINGS = ""
        udtAddData.strCOLOR = ""
        udtAddData.strROOT = ""
        Put #intFile, lngID, udtAddData
Done:
 Close #intFile
End Sub
