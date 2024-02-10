Attribute VB_Name = "mPrivProfFileTest"
' ----------------------------------------------------------------
' Standard Module mFsoTest: Test of all services of the module.
'
' ----------------------------------------------------------------
Private PP              As clsPrivProfFile

Private Property Let Test_Status(ByVal s As String)
    If s <> vbNullString Then
        Application.StatusBar = "Regression test " & ThisWorkbook.Name & " module 'mFso': " & s
    Else
        Application.StatusBar = vbNullString
    End If
End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mPrivProfFileTest." & sProc
End Function

Public Sub Test_00_Regression()
' ----------------------------------------------------------------------------
' All results are asserted and there is no intervention required for the whole
' test. When an assertion fails the test procedure will stop and indicates the
' problem with the called procedure. An execution trace is displayed at the
' end.
' ----------------------------------------------------------------------------
    Const PROC = "Test_00_Regression"

    On Error GoTo eh
    Dim sTestStatus As String
    
    '~~ Initialization (must be done prior the first BoP!)
    mTrc.FileName = "RegressionTest_clsPrivProfFile.ExecTrace.log"
    mTrc.Title = "Regression Test class module clsPrivProfFile"
    mTrc.NewFile
    mErH.Regression = True
    Set PP = New clsPrivProfFile ' the test runs with the default file name
    
    mBasic.BoP ErrSrc(PROC)
    sTestStatus = "clsPrivProfFile Regression Test: "

    mPrivProfFileTest.Test_89_FileName
    mPrivProfFileTest.Test_90_IsValidFileFullName
    mPrivProfFileTest.Test_91_PrivateProfile_SectionsNames
    mPrivProfFileTest.Test_92_PrivateProfile_ValueNames
    mPrivProfFileTest.Test_93_PrivateProfile_Value
    mPrivProfFileTest.Test_94_PrivateProfile_Values
    mPrivProfFileTest.Test_96_PrivateProfile_Entry_Exists
    mPrivProfFileTest.Test_97_PrivateProfile_SectionsCopy
    mPrivProfFileTest.Test_98_PrivateProfile_Reorg
    mPrivProfFileTest.Test_99_PrivateProfile_Reorg_WithNoFileProvided

xt: mTest.TestProc_RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    mErH.Regression = False
    mTrc.Dsply
    Set PP = Nothing
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_89_FileName()
    Const PROC = " Test_89_FileName"
    
    Dim s   As String
        
    mBasic.BoP ErrSrc(PROC)
    If PP Is Nothing Then Set PP = New clsPrivProfFile
    s = PP.FileName
    On Error Resume Next
    PP.FileName = s
    Debug.Assert Err.Number = 0
    Debug.Assert PP.FSo.FileExists(PP.FileName)
    mBasic.EoP ErrSrc(PROC)

End Sub

Public Sub Test_90_IsValidFileFullName()
    Const PROC = " Test_90_IsValidFileFullName"

    mBasic.BoP ErrSrc(PROC)
    If PP Is Nothing Then Set PP = New clsPrivProfFile
    Debug.Assert PP.IsValidFileName(ThisWorkbook.FullName)
    Debug.Assert Not PP.IsValidFileName("x")    ' missing :, missing \
    Debug.Assert Not PP.IsValidFileName("e:x")  ' missing \
    Debug.Assert Not PP.IsValidFileName("e:\x") ' missing extention
    Debug.Assert PP.IsValidFileName("e:\x.y")   ' complete with extention
    mBasic.EoP ErrSrc(PROC)

End Sub

Public Sub Test_91_PrivateProfile_SectionsNames()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_91_PrivateProfile_SectionsNames"
    
    On Error GoTo eh
    Dim sFileName   As String
    Dim dct         As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    If PP Is Nothing Then Set PP = New clsPrivProfFile
    sFileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues)
    PP.FileName = sFileName
    Set dct = PP.SectionNames()
    Debug.Assert dct.Count = 10
    Debug.Assert dct.Keys()(0) = mTest.TestProc_SectionName(1)
    Debug.Assert dct.Keys()(1) = mTest.TestProc_SectionName(2)
    Debug.Assert dct.Keys()(2) = mTest.TestProc_SectionName(3)

xt: mBasic.EoP ErrSrc(PROC)
    TestProc_RemoveTestFiles
    Set dct = Nothing
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_92_PrivateProfile_ValueNames()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_92_PrivateProfile_ValueNames"
    
    On Error GoTo eh
    Dim sFileName   As String
    Dim dct         As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    If PP Is Nothing Then Set PP = New clsPrivProfFile
    PP.FileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues)
    
    Set dct = PP.Values(v_section:=TestProc_SectionName(2))
    mBasic.EoP ErrSrc(PROC)
    Debug.Assert dct.Count = mTest.lTestValues
    Debug.Assert dct.Keys()(0) = TestProc_ValueName(2, 1)
    Debug.Assert dct.Keys()(1) = TestProc_ValueName(2, 2)
    Debug.Assert dct.Keys()(2) = TestProc_ValueName(2, 3)
       
xt: TestProc_RemoveTestFiles
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_93_PrivateProfile_Value()
' ----------------------------------------------------------------------------
' This test relies on the Value (Let) service.
' ----------------------------------------------------------------------------
    Const PROC = "Test_93_PrivateProfile_Value"
    
    On Error GoTo eh
    Dim cyValue     As Currency: cyValue = 12345.6789
    Dim cyResult    As Currency
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Test preparation
    Set PP = New clsPrivProfFile
    PP.FileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues)
            
    '~~ Test 1: Read non-existing value from a non-existing file
    mBasic.BoC "Value (using GetPrivateProfileString)"
    Debug.Assert PP.Value(v_value_name:="Any" _
                        , v_section:="Any" _
                         ) = vbNullString
    mBasic.EoC "Value (using GetPrivateProfileString)"
    
    '~~ Test 2: Read existing value
    mBasic.BoC "Value (using GetPrivateProfileString)"
    Debug.Assert mTest.TestProc_ValueString(3, 2) = PP.Value(v_value_name:=mTest.TestProc_ValueName(3, 2) _
                                                           , v_section:=mTest.TestProc_SectionName(3))
    mBasic.EoC "Value (using GetPrivateProfileString)"
    
    '~~ Test 3: Read non-existing without Lib functions
    mBasic.BoC "Value2 (not using GetPrivateProfileString)"
    Debug.Assert vbNullString = PP.Value2(v_value_name:="x" _
                                        , v_section:=mTest.TestProc_SectionName(3))
    mBasic.EoC "Value2 (not using GetPrivateProfileString)"
    mBasic.BoC "Value2 (not using GetPrivateProfileString)"
    Debug.Assert mTest.TestProc_ValueString(3, 2) = PP.Value2(v_value_name:=mTest.TestProc_ValueName(3, 2) _
                                                            , v_section:=mTest.TestProc_SectionName(3))
    mBasic.EoC "Value2 (not using GetPrivateProfileString)"
    
xt: TestProc_RemoveTestFiles
    Set PP = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_94_PrivateProfile_Values()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_94_PrivateProfile_Values"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim sFileName   As String
    
    mBasic.BoP ErrSrc(PROC)
    Set PP = New clsPrivProfFile
    PP.FileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues)

    '~~ Test 1: All values of one section
    Set dct = PP.Values(v_section:=TestProc_SectionName(2))
    '~~ Test 1: Assert the content of the result Dictionary
    Debug.Assert dct.Count = mTest.lTestValues
    Debug.Assert dct.Keys()(0) = TestProc_ValueName(2, 1)
    Debug.Assert dct.Keys()(1) = TestProc_ValueName(2, 2)
    Debug.Assert dct.Keys()(2) = TestProc_ValueName(2, 3)
    Debug.Assert dct.Items()(0) = TestProc_ValueString(2, 1)
    Debug.Assert dct.Items()(1) = TestProc_ValueString(2, 2)
    Debug.Assert dct.Items()(2) = TestProc_ValueString(2, 3)
    
    '~~ Test 2: No section provided
    Debug.Assert PP.Values().Count = 0

    '~~ Test 3: Section does not exist
    Debug.Assert PP.Values(v_file:=PP.FileName _
                         , v_section:="xxxxxxx").Count = 0

xt: TestProc_RemoveTestFiles
    Set dct = Nothing
    Set PP = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_96_PrivateProfile_Entry_Exists()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_96_PrivateProfile_Entry_Exists"

    On Error GoTo eh
    Dim sFileName   As String
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Test preparation
    Set PP = New clsPrivProfFile
    PP.FileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues)
       
    '~~ Section not exists
    Debug.Assert PP.SectionExists(s_file:=PP.FileName _
                                , s_section:=TestProc_SectionName(100) _
                                 ) = False
    '~~ Section exists
    Debug.Assert PP.SectionExists(s_file:=sFileName _
                                , s_section:=TestProc_SectionName(9) _
                                 ) = True
    '~~ Value-Name exists
    Debug.Assert PP.ValueExists(v_file:=sFileName _
                              , v_section:=TestProc_SectionName(7) _
                              , v_value_name:=TestProc_ValueName(7, 3) _
                                 ) = True
    '~~ Value-Name not exists
    Debug.Assert PP.ValueExists(v_file:=sFileName _
                              , v_section:=TestProc_SectionName(7) _
                              , v_value_name:=TestProc_ValueName(6, 3) _
                               ) = False
    
xt: TestProc_RemoveTestFiles
    Set PP = Nothing
    mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_97_PrivateProfile_SectionsCopy()
' ----------------------------------------------------------------------------
' This test relies on successfully tests:
' - Test_91_PrivateProfile_SectionsNames (PP.sectionNames)
' Iplicitely tested are:
' - PP.sections Get and Let
' ----------------------------------------------------------------------------
    Const PROC = "Test_97_PrivateProfile_SectionsCopy"
    
    On Error GoTo eh
    Dim sSourceFile     As String
    Dim sTargetFile     As String
    Dim sSectionName    As String
    Dim dct             As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    Set PP = New clsPrivProfFile
    
    '~~ Test 1a ------------------------------------
    '~~ Copy a specific section to a new target file
    sSourceFile = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues) ' prepare PrivateProfile test file
    PP.FileName = sSourceFile
    sTargetFile = ThisWorkbook.Path & "\Test\FsoTarget.dat"
    If PP.FSo.FileExists(sTargetFile) Then PP.FSo.DeleteFile sTargetFile
    
    sSectionName = PP.SectionNames(sSourceFile).Items()(0)
    PP.SectionsCopy s_source:=sSourceFile _
                  , s_target:=sTargetFile _
                  , s_sections:=sSectionName
    '~~ Assert result
    Set dct = PP.SectionNames(sTargetFile)
    Debug.Assert dct.Count = 1
    Debug.Assert dct.Keys()(0) = TestProc_SectionName(1)
    
    '~~ Test 1b ------------------------------------
    '~~ Copy a specific section to the target file of Test 1a
    PP.SectionsCopy s_source:=sSourceFile _
                  , s_target:=sTargetFile _
                  , s_sections:=TestProc_SectionName(8)
    '~~ Assert result
    Set dct = PP.SectionNames(sTargetFile)
    Debug.Assert dct.Count = 2
    Debug.Assert dct.Keys()(1) = TestProc_SectionName(8)
    PP.FSo.DeleteFile sSourceFile
    PP.FSo.DeleteFile sTargetFile
    
    '~~ Test 3 -------------------------------
    '~~ Copy all sections to a new target file (will be re-ordered ascending thereby)
    sSourceFile = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues) ' prepare PrivateProfile test file
    sTargetFile = ThisWorkbook.Path & "\Test\FsoTarget.dat"
    If PP.FSo.FileExists(sTargetFile) Then PP.FSo.DeleteFile sTargetFile
    PP.SectionsCopy s_source:=sSourceFile _
                  , s_target:=sTargetFile _
                  , s_sections:=PP.SectionNames(sSourceFile) _
                  , s_merge:=False
    '~~ Assert result
    Debug.Assert mFso.FileArry(sTargetFile)(0) = "[" & TestProc_SectionName(1) & "]"
    PP.FSo.DeleteFile sSourceFile
    PP.FSo.DeleteFile sTargetFile
            
xt: TestProc_RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_98_PrivateProfile_Reorg()
' ----------------------------------------------------------------------------
' Rearrange all sections and all names therein
' ----------------------------------------------------------------------------
    Const PROC = "Test_98_PrivateProfile_Reorg"
    
    On Error GoTo eh
    Dim sFileName   As String
    Dim vFile       As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Set PP = New clsPrivProfFile
    PP.FileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues) ' prepare PrivateProfile test file
    vFile = mFso.FileArry(PP.FileName)
    Debug.Assert vFile(0) = "[" & TestProc_SectionName(mTest.lTestSections) & "]"
    Debug.Assert vFile(1) = TestProc_ValueName(mTest.lTestSections, mTest.lTestValues) & "=" & TestProc_ValueString(mTest.lTestSections, mTest.lTestValues)
    
    PP.Reorg
    '~~ Assert result
    vFile = mFso.FileArry(PP.FileName)
    Debug.Assert vFile(0) = "[" & TestProc_SectionName(1) & "]"
    Debug.Assert vFile(1) = TestProc_ValueName(1, 1) & "=" & TestProc_ValueString(1, 1)
    
            
xt: TestProc_RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Set PP = Nothing
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_100_AllInOne()
    Const PROC = "Test_100_AllInOne"
    
    Dim i As Long
    Dim j As Long
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Test preparation
    Set PP = New clsPrivProfFile
    With PP
        .FileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues)
        .Reorg
    End With
    mBasic.EoP ErrSrc(PROC)
    Set PP = Nothing
    
End Sub

Public Sub Test_99_PrivateProfile_Reorg_WithNoFileProvided()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_99_PrivateProfile_NoFileProvided"
    
    On Error GoTo eh
    Dim sFileName   As String
    Dim vFile       As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Set PP = New clsPrivProfFile
    PP.FileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues) ' prepare PrivateProfile test file
    
    PP.Reorg
    '~~ Assert result
    vFile = mFso.FileArry(PP.FileName)
    Debug.Assert vFile(0) = "[" & TestProc_SectionName(1) & "]"
    Debug.Assert vFile(1) = TestProc_ValueName(1, 1) & "=" & TestProc_ValueString(1, 1)
              
xt: TestProc_RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub



