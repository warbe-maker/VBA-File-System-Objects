Attribute VB_Name = "mTest"
Option Explicit

Public Const lTestSections = 10
Public Const lTestValues = 15

Private Const SECTION_NAME = "Section_" ' for PrivateProfile services test
Private Const VALUE_NAME = "_Name_"     ' for PrivateProfile services test
Private Const VALUE_STRING = "-Value-"  ' for PrivateProfile services test

Private cllTestFiles    As Collection

Public Function TestProc_SectionName(ByVal l As Long) As String
    TestProc_SectionName = SECTION_NAME & Format(l, "00")
End Function

Public Function TestProc_ValueName(ByVal lS As Long, ByVal lV As Long) As String
    TestProc_ValueName = SECTION_NAME & Format(lS, "00") & VALUE_NAME & Format(lV, "00")
End Function

Public Function TestProc_ValueString(ByVal lS As Long, ByVal lV As Long) As String
    TestProc_ValueString = SECTION_NAME & Format(lS, "00") & VALUE_STRING & Format(lV, "00")
End Function

Public Function TestProc_PrivateProfile_File(Optional ByVal t_sections As Long = 10, _
                                             Optional ByVal t_values As Long = 15) As String
' ----------------------------------------------------------------------------
' Returns the name of a temporary file with n (t_sections) sections, each
' with m (t_values) values all in descending order. Each test file's name is
' saved to a Collection (cllTestFiles) allowing to delete them all at the end
' of the test.
' ----------------------------------------------------------------------------
    Const PROC = "TestProc_PrivateProfile_File"

    On Error GoTo eh
    Dim i           As Long
    Dim j           As Long
    Dim sFileName   As String

    mBasic.BoP ErrSrc(PROC)
    sFileName = ThisWorkbook.Path & "\Test\Fso.dat"
    If FSo.FileExists(sFileName) Then FSo.DeleteFile sFileName
    
    For i = t_sections To 1 Step -1
        For j = t_values To 1 Step -1
            mFso.PPvalue(pp_file:=sFileName _
                       , pp_section:=TestProc_SectionName(i) _
                       , pp_value_name:=TestProc_ValueName(i, j) _
                        ) = TestProc_ValueString(i, j)
        Next j
    Next i
    TestProc_PrivateProfile_File = sFileName
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sFileName

xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mFsoTest." & sProc
End Function

Public Sub TestProc_RemoveTestFiles()

    Dim v As Variant
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    For Each v In cllTestFiles
        If mFso.FSo.FileExists(v) Then
            Kill v
        End If
    Next v
    Set cllTestFiles = Nothing
    Set cllTestFiles = New Collection
    
End Sub

Public Function TestProc_TempFile() As String
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "TestProc_TempFile"

    On Error GoTo eh
    Dim sFileName   As String

    mBasic.BoP ErrSrc(PROC)
    sFileName = mFso.FileTemp(f_extension:=".dat")
    TestProc_TempFile = sFileName

    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sFileName

xt: mBasic.EoP ErrSrc(PROC)
    Exit Function

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function


