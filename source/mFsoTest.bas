Attribute VB_Name = "mFsoTest"
Option Explicit
' ----------------------------------------------------------------
' Standard Module mFsoTest: Test of all services of the module.
'
' ----------------------------------------------------------------
Private Const SECTION_NAME = "Section-" ' for PrivateProfile services test
Private Const VALUE_NAME = "-Name-"     ' for PrivateProfile services test
Private Const VALUE_STRING = "-Value-"  ' for PrivateProfile services test
    
Private cllTestFiles    As Collection

Private Property Get TestProc_SectionName(Optional ByVal l As Long)
    TestProc_SectionName = SECTION_NAME & Format(l, "00")
End Property

Private Property Get TestProc_ValueName(Optional ByVal lS As Long, Optional ByVal lV As Long)
    TestProc_ValueName = SECTION_NAME & Format(lS, "00") & VALUE_NAME & Format(lV, "00")
End Property

Private Property Get TestProc_ValueString(Optional ByVal lS As Long, Optional ByVal lV As Long)
    TestProc_ValueString = SECTION_NAME & Format(lS, "00") & VALUE_STRING & Format(lV, "00")
End Property

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

Private Sub BoC(ByVal boc_id As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(C)ode with id (boc_id) trace. Procedure to be copied as Private
' into any module potentially using the Common VBA Execution Trace Service. Has
' no effect when Conditional Compile Argument is 0 or not set at all.
' Note: The begin id (boc_id) has to be identical with the paired EoC statement.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ExecTrace = 1 Then
    mTrc.BoC boc_id, s
#End If
End Sub

Private Sub BoP(ByVal b_proc As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (B)egin-(o)f-(P)rocedure named (b_proc). Procedure to be copied as Private
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ErHComp = 1 Then
    mErH.BoP b_proc, s
#ElseIf ExecTrace = 1 Then
    mTrc.BoP b_proc, s
#End If
End Sub

Private Sub EoC(ByVal eoc_id As String, ParamArray b_arguments() As Variant)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(C)ode id (eoc_id) trace. Procedure to be copied as Private into
' any module potentially using the Common VBA Execution Trace Service. Has no
' effect when the Conditional Compile Argument is 0 or not set at all.
' Note: The end id (eoc_id) has to be identical with the paired BoC statement.
' ------------------------------------------------------------------------------
    Dim s As String: If UBound(b_arguments) >= 0 Then s = Join(b_arguments, ",")
#If ExecTrace = 1 Then
    mTrc.BoC eoc_id, s
#End If
End Sub

Private Sub EoP(ByVal e_proc As String, _
       Optional ByVal e_inf As String = vbNullString)
' ------------------------------------------------------------------------------
' (E)nd-(o)f-(P)rocedure named (e_proc). Procedure to be copied as Private Sub
' into any module potentially either using the Common VBA Error Service and/or
' the Common VBA Execution Trace Service. Has no effect when Conditional Compile
' Arguments are 0 or not set at all.
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    mErH.EoP e_proc
#ElseIf ExecTrace = 1 Then
    mTrc.EoP e_proc, e_inf
#End If
End Sub

Public Function ErrMsg(ByVal err_source As String, _
              Optional ByVal err_no As Long = 0, _
              Optional ByVal err_dscrptn As String = vbNullString, _
              Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service. Displays a debugging option button
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section when information is concatenated with
' the error message by two vertical bars (||).
'
' May be copied as Private Function into any module. Considers the Common VBA
' Message Service and the Common VBA Error Services as optional components.
' When neither is installed the error message is displayed by the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
' Note:  The above may seem to be a lot of code but will be a godsend in case
'        of an error!
'
' Uses:
' - AppErr For programmed application errors (Err.Raise AppErr(n), ....) to
'          turn tem into negative and in the error mesaage back into a positive
'          number.
' - ErrSrc To provide an unambigous procedure name - prefixed by the module name
'
' W. Rauschenberger Berlin, Nov 2021
'
' See:
' https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
#End If
    '~~ When neither of the Common Component is available in the VB-Project
    '~~ the error message is displayed by means of the VBA.MsgBox
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    '~~ Consider extra information is provided with the error description
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If err_dscrptn Like "*DAO*" _
            Or err_dscrptn Like "*ODBC*" _
            Or err_dscrptn Like "*Oracle*" _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & ErrDesc & vbLf & vbLf & "Source: " & vbLf & err_source & ErrAtLine
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mFsoTest." & sProc
End Function

Private Function TestProc_PrivateProfile_File(ByVal ts_sections As Long, _
                                              ByVal ts_values As Long) As String
' ----------------------------------------------------------------------------
' Returns the name of a temporary file with n (ts_sections) sections, each
' with m (ts_values) values all in descending order. Each test file's name is
' saved to a Collection (cllTestFiles) allowing to delete them all at the end
' of the test.
' ----------------------------------------------------------------------------
    Const PROC = "TestProc_PrivateProfile_File"
    
    On Error GoTo eh
    Dim i       As Long
    Dim j       As Long
    Dim sFile   As String
    
    BoP ErrSrc(PROC)
    sFile = mFso.FileTemp(tmp_extension:=".dat")
    For i = ts_sections To 1 Step -1
        For j = ts_values To 1 Step -1
            mFso.PPvalue(pp_file:=sFile _
                                 , pp_section:=TestProc_SectionName(i) _
                                 , pp_value_name:=TestProc_ValueName(i, j) _
                                  ) = TestProc_ValueString(i, j)
        Next j
    Next i
    TestProc_PrivateProfile_File = sFile
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sFile
    
xt: EoP ErrSrc(PROC)
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Sub TestProc_RemoveTestFiles()
    
    If mFso.Exists(ex_folder:=ThisWorkbook.Path, ex_file:="rad*.dat") Then
        Kill ThisWorkbook.Path & "\rad*.dat"
    End If
    
End Sub

Private Function TestProc_TempFile() As String
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "TestProc_TempFile"
    
    On Error GoTo eh
    Dim sFile   As String
    
    BoP ErrSrc(PROC)
    sFile = mFso.FileTemp(tmp_extension:=".dat")
    TestProc_TempFile = sFile
    
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sFile

xt: EoP ErrSrc(PROC)
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select

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
    
    '~~ Initialization of a new Trace Log File for this Regression test
    '~~ ! must be done prior the first BoP !
    mTrc.LogFile = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Regression Test.log")
    mTrc.LogTitle = "Regression Test module mFso"
        
    mErH.Regression = True
    
    BoP ErrSrc(PROC)
    sTestStatus = "mFso Regression Test: "

    mErH.Asserted AppErr(1) ' For the very last test on an error condition
    mFsoTest.Test_00_Regression_Common_Services
    mFsoTest.Test_00_Regression_PrivateProfile_Services
    
xt: TestProc_RemoveTestFiles
    EoP ErrSrc(PROC)
    mErH.Regression = False
    mTrc.Dsply
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_00_Regression_Common_Services()
' ----------------------------------------------------------------------------
' All results are asserted and there is no intervention required for the whole
' test. When an assertion fails the test procedure will stop and indicates the
' problem with the called procedure.
' ----------------------------------------------------------------------------
    Const PROC = "Test_00_Regression_Common_Services"

    On Error GoTo eh
    Dim sTestStatus As String
    
    sTestStatus = "mFso Regression-Other: "
    BoP ErrSrc(PROC)
    
    mErH.Asserted AppErr(1) ' For the very last test on an error condition
    mFsoTest.Test_01_File_Temp
    mFsoTest.Test_02_File_Exists
    mFsoTest.Test_07_File_Picked
    mFsoTest.Test_08_File_Txt_Let_Get
    mFsoTest.Test_09_File_Differs
    mFsoTest.Test_10_File_Arry_Get_Let
    mFsoTest.Test_11_File_Search
    mFsoTest.Test_12_IsValid_FileFolder_Name
    mFsoTest.Test_13_Folders_Test
    mFsoTest.Test_14_RenameSubFolders
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_00_Regression_PrivateProfile_Services()
' ----------------------------------------------------------------------------
' All results are asserted and there is no intervention required for the whole
' test. When an assertion fails the test procedure will stop and indicate the
' problem with the called procedure.
' ----------------------------------------------------------------------------
    Const PROC = "Test_00_Regression_PrivateProfile_Services"

    On Error GoTo eh
    Dim sTestStatus As String
    
    BoP ErrSrc(PROC)
    sTestStatus = "mFso Test_00_Regression_PrivateProfile_Services: "
    mErH.Asserted AppErr(1) ' For the very last test on an error condition
    
    mFsoTest.Test_91_PrivateProfile_SectionsNames
    mFsoTest.Test_92_PrivateProfile_ValueNames
    mFsoTest.Test_93_PrivateProfile_Value
    mFsoTest.Test_94_PrivateProfile_Values
    mFsoTest.Test_96_PrivateProfile_Entry_Exists
    mFsoTest.Test_97_PrivateProfile_SectionsCopy
    mFsoTest.Test_98_PrivateProfile_PPreorg
    mFsoTest.Test_99_PrivateProfile_PPreorg_WithNoFileProvided

xt: TestProc_RemoveTestFiles
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_01_File_Temp()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_01_File_Temp"

    Dim sTemp As String
    
    BoP ErrSrc(PROC)
    sTemp = mFso.FileTemp(tmp_path:=ThisWorkbook.Path)
    sTemp = mFso.FileTemp()
    EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_02_File_Exists()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_02_File_Exists"
    
    On Error GoTo eh
    Dim cll     As Collection
    Dim fso     As New FileSystemObject
    Dim sFile   As String
    
    BoP ErrSrc(PROC)
    
    '~~ Folder exists
    Debug.Assert mFso.Exists(ex_folder:=ThisWorkbook.Path & "x") = False
    Debug.Assert mFso.Exists(ex_folder:=ThisWorkbook.Path) = True
    
    '~~ File exists
    Debug.Assert mFso.Exists(ex_file:=ThisWorkbook.FullName & "x") = False
    Debug.Assert mFso.Exists(ex_file:=ThisWorkbook.FullName) = True

    '~~ Section exists
    sFile = TestProc_PrivateProfile_File(ts_sections:=3, ts_values:=3)
    Debug.Assert mFso.Exists(ex_file:=sFile _
                            , ex_section:=TestProc_SectionName(2) & "x" _
                             ) = False
    Debug.Assert mFso.Exists(ex_file:=sFile _
                            , ex_section:=TestProc_SectionName(2) _
                             ) = True
    
    '~~ Value-Name exists
    Debug.Assert mFso.Exists(ex_file:=sFile _
                            , ex_section:=TestProc_SectionName(2) _
                            , ex_value_name:=TestProc_ValueName(2, 2) & "x" _
                             ) = False
    Debug.Assert mFso.Exists(ex_file:=sFile _
                            , ex_section:=TestProc_SectionName(2) _
                            , ex_value_name:=TestProc_ValueName(2, 2) _
                             ) = True

    '~~ File by wildcard, in any sub-folder, exactly one
    Debug.Assert mFso.Exists(ex_folder:=ThisWorkbook.Path _
                            , ex_file:="*.xl*" _
                            , ex_result_files:=cll) = True
    Debug.Assert cll.Count = 1
    Debug.Assert cll(1).Path = ThisWorkbook.FullName
            
    '~~ File by wildcard, in any sub-folder, more than one
    Debug.Assert mFso.Exists(ex_folder:=ThisWorkbook.Path _
                            , ex_file:="fMsg.fr*" _
                            , ex_result_files:=cll) = True
    Debug.Assert cll.Count = 2
    Debug.Assert cll(1).Name = "fMsg.frm"
    Debug.Assert cll(2).Name = "fMsg.frx"
                        
xt: TestProc_RemoveTestFiles
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_03_FilePathSplit()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_03_FilePathSplit"
    
    On Error GoTo eh
    Dim p   As String   ' folder path
    Dim f   As String   ' file
    
    BoP ErrSrc(PROC)
    p = ThisWorkbook.FullName
    mFso.FilePathSplit p, f
    Debug.Assert p = ThisWorkbook.Path & "\"
    Debug.Assert f = ThisWorkbook.Name

    p = ThisWorkbook.Name
    mFso.FilePathSplit p, f
    Debug.Assert p = vbNullString
    Debug.Assert f = ThisWorkbook.Name

    p = ThisWorkbook.Path
    mFso.FilePathSplit p, f
    Debug.Assert p = ThisWorkbook.Path
    Debug.Assert f = vbNullString
    
    p = ThisWorkbook.FullName & "x"
    mFso.FilePathSplit p, f
    Debug.Assert p = vbNullString
    Debug.Assert f = ThisWorkbook.FullName & "x" ' any last element with a . is interpreted as file name
    
    p = ThisWorkbook.Path & "x"
    mFso.FilePathSplit p, f
    Debug.Assert p = ThisWorkbook.Path & "x"
    Debug.Assert f = vbNullString

xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_07_File_Picked()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_07_File_Picked"
    
    On Error GoTo eh
    Dim fl As File

    BoP ErrSrc(PROC)
    Test_Status = ErrSrc(PROC)
    
    '~~ Test 1: A file is picked
    If mFso.FilePicked(p_init_path:=ThisWorkbook.Path _
                     , p_filters:="Excel file, *.xl*; All files, *.*" _
                     , p_title:="Test 1: Select the Excel Workbook in this folder (folder preselected by filter)" _
                     , p_file:=fl _
                      ) = True Then
        Debug.Assert fl.Path = ThisWorkbook.FullName
    Else
        Debug.Assert fl Is Nothing
    End If
    
    '~~ Test 2: No file picked
    Debug.Assert Not mFso.FilePicked(p_init_path:=ThisWorkbook.Path _
                                   , p_filters:="Excel file, *.xl*; All files, *.*" _
                                   , p_title:="Test 2: No file picked (just  t e r m i n a t e  the dialog)" _
                                   , p_file:=fl _
                                    )
    Debug.Assert fl Is Nothing
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_08_File_Txt_Let_Get()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_08_File_Txt_Let_Get"
    
    On Error GoTo eh
    Dim sFl     As String
    Dim sTest   As String
    Dim sResult As String
    Dim sSplit  As String
    Dim fso     As New FileSystemObject
    Dim oFl     As File
    
    Test_Status = ErrSrc(PROC)
    BoP ErrSrc(PROC)
    
    '~~ Test 1: Write one recod
    sFl = mFso.FileTemp()
    sTest = "My string"
    mFso.FileTxt(ft_file:=sFl _
            , ft_append:=False _
             ) = sTest
    sResult = mFso.FileTxt(ft_file:=sFl, ft_split:=sSplit)
    Debug.Assert Split(sResult, sSplit)(0) = sTest
    fso.DeleteFile sFl
    
    '~~ Test 2: Empty file
    sFl = mFso.FileTemp()
    sTest = vbNullString
    mFso.FileTxt(ft_file:=sFl, ft_append:=False) = sTest
    sResult = mFso.FileTxt(ft_file:=sFl, ft_split:=sSplit)
    Debug.Assert sResult = vbNullString
    fso.DeleteFile sFl

    '~~ Test 3: Append
    sFl = mFso.FileTemp()
    mFso.FileTxt(ft_file:=sFl, ft_append:=False) = "AAA" & vbCrLf & "BBB"
    mFso.FileTxt(ft_file:=sFl, ft_append:=True) = "CCC"
    sResult = mFso.FileTxt(ft_file:=sFl, ft_split:=sSplit)
    Debug.Assert Split(sResult, sSplit)(0) = "AAA"
    Debug.Assert Split(sResult, sSplit)(1) = "BBB"
    Debug.Assert Split(sResult, sSplit)(2) = "CCC"
    fso.DeleteFile sFl

    '~~ Test 4: Write with append and read with file as object
    sFl = mFso.FileTemp()
    fso.CreateTextFile FileName:=sFl
    Set oFl = fso.GetFile(sFl)
    sFl = oFl.Path
    mFso.FileTxt(ft_file:=oFl, ft_append:=False) = "AAA" & vbCrLf & "BBB"
    mFso.FileTxt(ft_file:=oFl, ft_append:=True) = "CCC"
    sResult = mFso.FileTxt(ft_file:=oFl, ft_split:=sSplit)
    Debug.Assert Split(sResult, sSplit)(0) = "AAA"
    Debug.Assert Split(sResult, sSplit)(1) = "BBB"
    Debug.Assert Split(sResult, sSplit)(2) = "CCC"
    fso.DeleteFile sFl

xt: Set fso = Nothing
    EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_09_File_Differs()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_09_File_Differs"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim f1      As File
    Dim f2      As File
    Dim dctDiff As Dictionary
    Dim sF1     As String
    Dim sF2     As String

    BoP ErrSrc(PROC)
    Test_Status = ErrSrc(PROC)
    
    sF1 = mFso.FileTemp
    sF2 = mFso.FileTemp

    BoP ErrSrc(PROC)
    ' Prepare
    mFso.FileTxt(ft_file:=sF1, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    mFso.FileTxt(ft_file:=sF2, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    Set f1 = fso.GetFile(sF1)
    Set f2 = fso.GetFile(sF2)

    ' Test 1: Differs.Count = 0
    Set dctDiff = mFso.FileDiffers(fd_file1:=f1 _
                              , fd_file2:=f2 _
                              , fd_ignore_empty_records:=True _
                              , fd_stop_after:=2 _
                               )
    Debug.Assert dctDiff.Count = 0

    ' Test 2: Differs.Count = 1
    mFso.FileTxt(ft_file:=sF1, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    mFso.FileTxt(ft_file:=sF2, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C" & vbCrLf & "D"
    Set f1 = fso.GetFile(sF1)
    Set f2 = fso.GetFile(sF2)
    
    Set dctDiff = mFso.FileDiffers(fd_file1:=f1 _
                              , fd_file2:=f2 _
                              , fd_ignore_empty_records:=True _
                              , fd_stop_after:=2 _
                               )
    Debug.Assert dctDiff.Count = 1
    
    ' Test 3: Differs.Count = 1
    mFso.FileTxt(ft_file:=sF1, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C" & vbCrLf & "D"
    mFso.FileTxt(ft_file:=sF2, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    Set f1 = fso.GetFile(sF1)
    Set f2 = fso.GetFile(sF2)
    
    Set dctDiff = mFso.FileDiffers(fd_file1:=f1 _
                              , fd_file2:=f2 _
                              , fd_ignore_empty_records:=True _
                              , fd_stop_after:=2 _
                               )
    Debug.Assert dctDiff.Count = 1
    
    ' Test 4: Differs.Count = 1
    mFso.FileTxt(ft_file:=sF1, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    mFso.FileTxt(ft_file:=sF2, ft_append:=False) = "A" & vbCrLf & "X" & vbCrLf & "C"
    Set f1 = fso.GetFile(sF1)
    Set f2 = fso.GetFile(sF2)
    
    Set dctDiff = mFso.FileDiffers(fd_file1:=f1 _
                              , fd_file2:=f2 _
                              , fd_ignore_empty_records:=True _
                              , fd_stop_after:=2 _
                               )
    Debug.Assert dctDiff.Count = 1

xt: EoP ErrSrc(PROC)
    If fso.FileExists(sF1) Then fso.DeleteFile (sF1)
    If fso.FileExists(sF2) Then fso.DeleteFile (sF2)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

  
Public Sub Test_09_File_Differs_False()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_09_File_Differs"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim sFile   As String
    Dim f1      As File
    Dim f2      As File
    Dim dctDiff As Dictionary
    
    Test_Status = ErrSrc(PROC)
    ' Prepare
    sFile = fso.GetFolder(ThisWorkbook.Path).ParentFolder.Path & "\Common-Components\mFso.bas"
    Set f1 = fso.GetFile(sFile)
    Set f2 = fso.GetFile(sFile)
    BoP ErrSrc(PROC), "fd_file1 = ", f1.Name, "fd_file2 = ", f2.Name
    
    ' Test
    Set dctDiff = mFso.FileDiffers(fd_file1:=f1, fd_file2:=f2, fd_ignore_empty_records:=True)
    Debug.Assert dctDiff.Count = 0

xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

  
Public Sub Test_10_File_Arry_Get_Let()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_10_File_Arry_Get_Let"
    
    On Error GoTo eh
    Dim sFile1      As String
    Dim sFile2      As String
    Dim lInclEmpty  As Long
    Dim lEmpty1     As Long
    Dim lExclEmpty  As Long
    Dim lEmpty2     As Long
    Dim fso         As New FileSystemObject
    Dim a           As Variant
    Dim v           As Variant
    
    BoP ErrSrc(PROC)
    Test_Status = ErrSrc(PROC)
    
    sFile1 = fso.GetFolder(ThisWorkbook.Path).ParentFolder.Path & "Common-Components\mFso.bas"
    sFile2 = fso.GetFolder(ThisWorkbook.Path).ParentFolder.Path & "Common-Components\mFso.bas"
    
    sFile1 = mFso.FileTemp()
    sFile2 = mFso.FileTemp()
    
    '~~ Write to lines to sFile1
    mFso.FileTxt(sFile1) = "xxx" & vbCrLf & "" & "yyy"
    
    '~~ Get the two lines as Array
    a = mFso.FileArry(fa_file:=sFile1 _
                 , fa_split:=vbCrLf _
                  )
    Debug.Assert a(LBound(a)) = "xxx"
    Debug.Assert a(UBound(a)) = "yyy"

    '~~ Write array to file-2
    mFso.FileArry(fa_file:=sFile2 _
             , fa_split:=vbCrLf _
              ) = a
    Debug.Assert mFso.FileDiffers(fso.GetFile(sFile1), fso.GetFile(sFile2)).Count = 0

    '~~ Count empty records when array contains all text lines
    a = mFso.FileArry(fa_file:=sFile1, fa_excl_empty_lines:=False)
    lInclEmpty = UBound(a) + 1
    lEmpty1 = 0
    For Each v In a
        If VBA.Trim$(v) = vbNullString Then lEmpty1 = lEmpty1 + 1
        If VBA.Len(Trim$(v)) = 0 Then lEmpty2 = lEmpty2 + 1
    Next v
    
    '~~ Count empty records
    a = mFso.FileArry(fa_file:=sFile1, fa_excl_empty_lines:=True)
    lExclEmpty = UBound(a) + 1
    Debug.Assert lExclEmpty = lInclEmpty - lEmpty1
    
xt: With fso
        .DeleteFile sFile1
        If .FileExists(sFile2) Then .DeleteFile sFile2
    End With
    Set fso = Nothing
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_11_File_Search()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_11_File_Search"
    
    On Error GoTo eh
    Dim cll As Collection
    Dim fso As New FileSystemObject
    
    BoP ErrSrc(PROC)
    Test_Status = ErrSrc(PROC)
    
    '~~ Test 1: Including subfolders, several files found
    Set cll = mFso.FilesSearch(fs_root:=fso.GetFolder(ThisWorkbook.Path).ParentFolder.Path & "\Common-Components\" _
                             , fs_mask:="*.bas*" _
                             , fs_stop_after:=5 _
                              )
    Debug.Assert cll.Count > 2

    '~~ Test 2: Not including subfolders, no files found
    Set cll = mFso.FilesSearch(fs_root:="e:\Ablage\Excel VBA\DevAndTest\Common" _
                             , fs_mask:="*CompMan*.frx" _
                             , fs_stop_after:=5 _
                             , fs_in_subfolders:=False _
                              )
    Debug.Assert cll.Count = 0

xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_12_IsValid_FileFolder_Name()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_12_IsValid_FileFolder_Name"
    
    BoP ErrSrc(PROC)
    '~~ Test 1: Valid Folder Name
    Debug.Assert mFso.FolderIsValidName(ThisWorkbook.Path) = True        ' a valid folder is a valid file name as well
    Debug.Assert mFso.FolderIsValidName(ThisWorkbook.FullName) = True
    Debug.Assert mFso.FolderIsValidName(ThisWorkbook.Name) = False
    Debug.Assert mFso.FolderIsValidName("c:\LP?1") = False

    '~~ Test 2: Valid File Name
    Debug.Assert mFso.FileIsValidName(ThisWorkbook.Name) = True
    Debug.Assert mFso.FileIsValidName(ThisWorkbook.Name & "?") = False
    EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_13_Folders_Test()
    Const PROC = "Test_13_Folders_Test"
    
    Dim TestFolder As String
    Dim fso As New FileSystemObject
    
    TestFolder = fso.GetFolder(ThisWorkbook.Path).ParentFolder.Path
    
    Dim v       As Variant
    Dim cll     As Collection
    Dim s       As String
    Dim sStart  As String
    
    Set cll = Folders("c:\XXXX", True, sStart)
    s = "1. Test: Folders in a provided non-existing folder ('" & sStart & "')"
    Debug.Print vbLf & s
    Debug.Print String(Len(s), "-")
    Debug.Assert cll.Count = 0
    
    Set cll = Folders(TestFolder, , sStart)
    s = "2. Test: Folders in the provided folder '" & sStart & "' (without sub-folders):"
    Debug.Print vbLf & s
    Debug.Print String(Len(s), "-")
    For Each v In cll
        Debug.Print v.Path
    Next v

    Set cll = Folders(TestFolder, True, sStart)
    s = "3. Test: Folders in the provided folder '" & sStart & "' (including sub-folders):"
    Debug.Print vbLf & s
    Debug.Print String(Len(s), "-")
    For Each v In cll
        Debug.Print v.Path
    Next v

    Set cll = Folders(, True, sStart)
    s = "4. Test: Folders in the manually selected folder '" & sStart & "' (including sub-folders):"
    Debug.Print vbLf & s
    Debug.Print String(Len(s), "-")
    For Each v In cll
        Debug.Print v.Path
    Next v
        
    Set fso = Nothing
End Sub

Public Sub Test_14_RenameSubFolders()
    Const PROC = "Test_14_RenameSubFolders"
    
    Dim cllRenamed  As Collection
    
    BoP ErrSrc(PROC)
    mFso.RenameSubFolders ThisWorkbook.Path & "\Test", "Sub_2_2", "Sub_2_2_renamed", cllRenamed
    Debug.Assert cllRenamed.Count = 1
    Debug.Assert cllRenamed(1).Path = ThisWorkbook.Path & "\Test\Sub_2\Sub_2_2_renamed"

    mFso.RenameSubFolders ThisWorkbook.Path & "\Test", "Sub_2_2_renamed", "Sub_2_2", cllRenamed
    Debug.Assert cllRenamed.Count = 1
    Debug.Assert cllRenamed(1).Path = ThisWorkbook.Path & "\Test\Sub_2\Sub_2_2"
    EoP ErrSrc(PROC)

End Sub

Public Sub Test_91_PrivateProfile_SectionsNames()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_91_PrivateProfile_SectionsNames"
    
    On Error GoTo eh
    Dim sFile   As String
    Dim fso     As New FileSystemObject
    Dim dct     As Dictionary
    
    BoP ErrSrc(PROC)
    sFile = TestProc_PrivateProfile_File(ts_sections:=3, ts_values:=3)
    
    Set dct = mFso.PPsectionNames(sFile)
    Debug.Assert dct.Count = 3
    Debug.Assert dct.Items()(0) = TestProc_SectionName(1)
    Debug.Assert dct.Items()(1) = TestProc_SectionName(2)
    Debug.Assert dct.Items()(2) = TestProc_SectionName(3)

xt: EoP ErrSrc(PROC)
    TestProc_RemoveTestFiles
    Set fso = Nothing
    Set dct = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
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
    Dim sFile   As String
    Dim dct     As Dictionary
    Dim fso     As New FileSystemObject
    
    BoP ErrSrc(PROC)
    sFile = TestProc_PrivateProfile_File(ts_sections:=5, ts_values:=3)
    
    Set dct = mFso.PPvalues(pp_file:=sFile, pp_section:=TestProc_SectionName(2))
    EoP ErrSrc(PROC)
    Debug.Assert dct.Count = 3
    Debug.Assert dct.Keys()(0) = TestProc_ValueName(2, 1)
    Debug.Assert dct.Keys()(1) = TestProc_ValueName(2, 2)
    Debug.Assert dct.Keys()(2) = TestProc_ValueName(2, 3)
       
xt: TestProc_RemoveTestFiles
    Set fso = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
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
    Dim fso         As New FileSystemObject
    Dim sFile       As String
    Dim cyValue     As Currency: cyValue = 12345.6789
    Dim cyResult    As Currency
    
    BoP ErrSrc(PROC)
    '~~ Test preparation
    sFile = TestProc_TempFile
            
    '~~ Test 1: Read non-existing value from a non-existing file
    Debug.Assert mFso.PPvalue(pp_file:=sFile _
                           , pp_section:="Any" _
                           , pp_value_name:="Any" _
                            ) = vbNullString
    
    '~~ Test 2: Write values
    mFso.PPvalue(pp_file:=sFile, pp_section:=TestProc_SectionName(1), pp_value_name:=TestProc_ValueName(1, 1)) = TestProc_ValueString(1, 1)
    mFso.PPvalue(pp_file:=sFile, pp_section:=TestProc_SectionName(1), pp_value_name:=TestProc_ValueName(1, 2)) = TestProc_ValueString(1, 2)
    mFso.PPvalue(pp_file:=sFile, pp_section:=TestProc_SectionName(2), pp_value_name:=TestProc_ValueName(2, 1)) = TestProc_ValueString(2, 1)
    mFso.PPvalue(pp_file:=sFile, pp_section:=TestProc_SectionName(2), pp_value_name:=TestProc_ValueName(2, 2)) = cyValue
    
    '~~ Test 2: Assert written values
    Debug.Assert mFso.PPvalue(pp_file:=sFile, pp_section:=TestProc_SectionName(1), pp_value_name:=TestProc_ValueName(1, 1)) = TestProc_ValueString(1, 1)
    Debug.Assert mFso.PPvalue(pp_file:=sFile, pp_section:=TestProc_SectionName(1), pp_value_name:=TestProc_ValueName(1, 2)) = TestProc_ValueString(1, 2)
    Debug.Assert mFso.PPvalue(pp_file:=sFile, pp_section:=TestProc_SectionName(2), pp_value_name:=TestProc_ValueName(2, 1)) = TestProc_ValueString(2, 1)
    cyResult = mFso.PPvalue(pp_file:=sFile, pp_section:=TestProc_SectionName(2), pp_value_name:=TestProc_ValueName(2, 2))
    Debug.Assert cyResult = cyValue
    Debug.Assert VarType(cyResult) = vbCurrency
    
xt: TestProc_RemoveTestFiles
    Set fso = Nothing
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
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
    Dim sFile       As String
    Dim fso         As New FileSystemObject
    
    BoP ErrSrc(PROC)
    sFile = TestProc_PrivateProfile_File(ts_sections:=3 _
                               , ts_values:=3 _
                                )

    '~~ Test 1: All values of one section
    Set dct = mFso.PPvalues(pp_file:=sFile _
                         , pp_section:=TestProc_SectionName(2) _
                          )
    '~~ Test 1: Assert the content of the result Dictionary
    Debug.Assert dct.Count = 3
    Debug.Assert dct.Keys()(0) = TestProc_ValueName(2, 1)
    Debug.Assert dct.Keys()(1) = TestProc_ValueName(2, 2)
    Debug.Assert dct.Keys()(2) = TestProc_ValueName(2, 3)
    Debug.Assert dct.Items()(0) = TestProc_ValueString(2, 1)
    Debug.Assert dct.Items()(1) = TestProc_ValueString(2, 2)
    Debug.Assert dct.Items()(2) = TestProc_ValueString(2, 3)
    
    '~~ Test 2: No section provided
    Debug.Assert mFso.PPvalues(sFile, vbNullString).Count = 0

    '~~ Test 3: Section does not exist
    Debug.Assert mFso.PPvalues(sFile, "xxxxxxx").Count = 0

xt: TestProc_RemoveTestFiles
    Set dct = Nothing
    Set fso = Nothing
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
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
    Dim sFile   As String
    
    BoP ErrSrc(PROC)
    '~~ Test preparation
    sFile = TestProc_PrivateProfile_File(ts_sections:=10, ts_values:=3)
       
    '~~ Section not exists
    Debug.Assert mFso.PPsectionExists(pp_file:=sFile _
                                              , pp_section:=TestProc_SectionName(100) _
                                               ) = False
    '~~ Section exists
    Debug.Assert mFso.PPsectionExists(pp_file:=sFile _
                                              , pp_section:=TestProc_SectionName(9) _
                                              ) = True
    '~~ Value-Name exists
    Debug.Assert mFso.PPvalueExists(pp_file:=sFile _
                                , pp_section:=TestProc_SectionName(7) _
                                , pp_value_name:=TestProc_ValueName(7, 3) _
                                 ) = True
    '~~ Value-Name not exists
    Debug.Assert mFso.PPvalueExists(pp_file:=sFile _
                                , pp_section:=TestProc_SectionName(7) _
                                , pp_value_name:=TestProc_ValueName(6, 3) _
                                 ) = False
    
xt: TestProc_RemoveTestFiles
    EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_97_PrivateProfile_SectionsCopy()
' ----------------------------------------------------------------------------
' This test relies on successfully tests:
' - Test_91_PrivateProfile_SectionsNames (mFso.PPsectionNames)
' Iplicitely tested are:
' - mFso.PPsections Get and Let
' ----------------------------------------------------------------------------
    Const PROC = "Test_97_PrivateProfile_SectionsCopy"
    
    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim SourceFile      As String
    Dim TargetFile      As String
    Dim sSectionName    As String
    Dim dct             As Dictionary
    
    BoP ErrSrc(PROC)
    
    '~~ Test 1a ------------------------------------
    '~~ Copy a specific section to a new target file
    SourceFile = TestProc_PrivateProfile_File(ts_sections:=10, ts_values:=10) ' prepare PrivateProfile test file
    TargetFile = mFso.FileTemp(tmp_extension:=".dat")
    sSectionName = mFso.PPsectionNames(SourceFile).Items()(0)
    mFso.PPsectionsCopy pp_source:=SourceFile _
                                , pp_target:=TargetFile _
                                , pp_sections:=sSectionName
    '~~ Assert result
    Set dct = mFso.PPsectionNames(TargetFile)
    Debug.Assert dct.Count = 1
    Debug.Assert dct.Keys()(0) = TestProc_SectionName(1)
    
    '~~ Test 1b ------------------------------------
    '~~ Copy a specific section to the target file of Test 1a
    mFso.PPsectionsCopy pp_source:=SourceFile _
                                , pp_target:=TargetFile _
                                , pp_sections:=TestProc_SectionName(8)
    '~~ Assert result
    Set dct = mFso.PPsectionNames(TargetFile)
    Debug.Assert dct.Count = 2
    Debug.Assert dct.Keys()(1) = TestProc_SectionName(8)
    fso.DeleteFile SourceFile
    fso.DeleteFile TargetFile
    
    '~~ Test 3 -------------------------------
    '~~ Copy all sections to a new target file (will be re-ordered ascending thereby)
    SourceFile = TestProc_PrivateProfile_File(ts_sections:=10, ts_values:=10) ' prepare PrivateProfile test file
    TargetFile = mFso.FileTemp(tmp_extension:=".dat")
    mFso.PPsectionsCopy pp_source:=SourceFile _
                                , pp_target:=TargetFile _
                                , pp_sections:=mFso.PPsectionNames(SourceFile) _
                                , pp_merge:=False
    '~~ Assert result
    Debug.Assert mFso.FileArry(TargetFile)(0) = "[" & TestProc_SectionName(1) & "]"
    fso.DeleteFile SourceFile
    fso.DeleteFile TargetFile
            
xt: TestProc_RemoveTestFiles
    Set fso = Nothing
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_98_PrivateProfile_PPreorg()
' ----------------------------------------------------------------------------
' Rearrange all sections and all names therein
' ----------------------------------------------------------------------------
    Const PROC = "Test_98_PrivateProfile_PPreorg"
    
    On Error GoTo eh
    Dim sFile   As String
    Dim vFile   As Variant
    
    BoP ErrSrc(PROC)
    
    sFile = TestProc_PrivateProfile_File(ts_sections:=10, ts_values:=10) ' prepare PrivateProfile test file
    vFile = mFso.FileArry(sFile)
    Debug.Assert vFile(0) = "[" & TestProc_SectionName(10) & "]"
    Debug.Assert vFile(1) = TestProc_ValueName(10, 10) & "=" & TestProc_ValueString(10, 10)
    
    mFso.PPreorg sFile
    '~~ Assert result
    vFile = mFso.FileArry(sFile)
    Debug.Assert vFile(0) = "[" & TestProc_SectionName(1) & "]"
    Debug.Assert vFile(1) = TestProc_ValueName(1, 1) & "=" & TestProc_ValueString(1, 1)
    
            
xt: TestProc_RemoveTestFiles
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_99_PrivateProfile_PPreorg_WithNoFileProvided()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_99_PrivateProfile_NoFileProvided"
    
    On Error GoTo eh
    Dim sFile   As String
    Dim vFile   As Variant
    
    BoP ErrSrc(PROC)
    sFile = TestProc_PrivateProfile_File(ts_sections:=10, ts_values:=10) ' prepare PrivateProfile test file
    
    mFso.PPreorg sFile
    '~~ Assert result
    vFile = mFso.FileArry(sFile)
    Debug.Assert vFile(0) = "[" & TestProc_SectionName(1) & "]"
    Debug.Assert vFile(1) = TestProc_ValueName(1, 1) & "=" & TestProc_ValueString(1, 1)
              
xt: TestProc_RemoveTestFiles
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

