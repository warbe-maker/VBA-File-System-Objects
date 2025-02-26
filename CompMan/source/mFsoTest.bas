Attribute VB_Name = "mFsoTest"
Option Explicit
' ----------------------------------------------------------------
' Standard Module mFsoTest: Test of all services of the module.
'
' Uses: clsTestAid and clsLog
'
' W. Rauschenberger, Berlin Apr 2024
' ----------------------------------------------------------------
Private Const SECTION_NAME  As String = "Section-"  ' for PrivateProfile services test
Private Const VALUE_NAME    As String = "-Name-"    ' for PrivateProfile services test
Private Const VALUE_STRING  As String = "-Value-"   ' for PrivateProfile services test
Private Tests               As clsTestAid
Private bRegression         As Boolean

'~~ When the a copy of VarTrans and all its depending Functions is used in a VB-Project,
'~~ e.g. to provide an independency of a module, this will have to be copied in a
'~~ Standard module of the respective VB-Project. VarTrans procedure.
'~~ Originates in Common Component mVarTrans.
Public Enum VarTransAs
    enAsArray
    enAsCollection
    enAsDictionary
    enAsFile
    enAsString
End Enum
'~~ --------------------------------------------------------------------

Private Property Get AsArray() As Long:  AsArray = 1:        End Property

Private Property Get AsCollection():     AsCollection = 2:   End Property

Private Property Get AsDictionary():     AsDictionary = 3:   End Property

Private Property Get AsFile():           AsFile = 4:         End Property

Private Property Get AsString():         AsString = 5:       End Property

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

Private Sub BoP(ByVal b_proc As String, _
       Optional ByVal b_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mErH Then          ' serves the mTrc/clsTrc when installed and active
    mErH.BoP b_proc, b_args
#ElseIf XcTrc_clsTrc Then ' when only clsTrc is installed and active
    If Trc Is Nothing Then Set Trc = New clsTrc
    Trc.BoP b_proc, b_args
#ElseIf XcTrc_mTrc Then   ' when only mTrc is installed and activate
    mTrc.BoP b_proc, b_args
#End If
End Sub

    
Private Sub EoP(ByVal e_proc As String, _
       Optional ByVal e_args As String = vbNullString)
' ------------------------------------------------------------------------------
' Common 'Begin of Procedure' interface serving the 'Common VBA Error Services'
' and - if not installed/activated the 'Common VBA Execution Trace Service'.
' Obligatory copy Private for any VB-Component using the service but not having
' the mBasic common component installed.
' ------------------------------------------------------------------------------
#If mErH = 1 Then          ' serves the mTrc/clsTrc when installed and active
    mErH.EoP e_proc, e_args
#ElseIf clsTrc = 1 Then ' when only clsTrc is installed and active
    Trc.EoP e_proc, e_args
#ElseIf mTrc = 1 Then   ' when only mTrc is installed and activate
    mTrc.EoP e_proc, e_args
#End If
End Sub

Public Function ErrMsg(ByVal err_source As String, _
              Optional ByVal err_no As Long = 0, _
              Optional ByVal err_dscrptn As String = vbNullString, _
              Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service which displays:
' - a debugging option button
' - an "About:" section when the err_dscrptn has an additional string
'   concatenated by two vertical bars (||)
' - the error message either by means of the Common VBA Message Service
'   (fMsg/mMsg) when installed (indicated by Cond. Comp. Arg. `mMsg = 1` or by
'   means of the VBA.MsgBox in case not.
'
' Uses: AppErr  For programmed application errors (Err.Raise AppErr(n), ....)
'               to turn them into a negative and in the error message back into
'               its origin positive number.
'
' W. Rauschenberger Berlin, Jan 2024
' See: https://github.com/warbe-maker/VBA-Error
' ------------------------------------------------------------------------------
#If mErH = 1 Then
    '~~ When Common VBA Error Services (mErH) is availabel in the VB-Project
    '~~ (which includes the mMsg component) the mErh.ErrMsg service is invoked.
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
#ElseIf mMsg = 1 Then
    '~~ When (only) the Common Message Service (mMsg, fMsg) is available in the
    '~~ VB-Project, mMsg.ErrMsg is invoked for the display of the error message.
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line): GoTo xt
    GoTo xt
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
    '~~ About
    ErrDesc = err_dscrptn
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    End If
    '~~ Type of error
    If err_no < 0 Then
        ErrType = "Application Error ": ErrNo = AppErr(err_no)
    Else
        ErrType = "VB Runtime Error ":  ErrNo = err_no
        If err_dscrptn Like "*DAO*" _
        Or err_dscrptn Like "*ODBC*" _
        Or err_dscrptn Like "*Oracle*" _
        Then ErrType = "Database Error "
    End If
    
    '~~ Title
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")
    '~~ Description
    ErrText = "Error: " & vbLf & ErrDesc
    '~~ About
    If ErrAbout <> vbNullString Then ErrText = ErrText & vbLf & vbLf & "About: " & vbLf & ErrAbout
    
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & "Debugging:" & vbLf & "Yes    = Resume Error Line" & vbLf & "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle, Prompt:=ErrText, Buttons:=ErrBttns)
xt:
End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mFsoTest." & sProc
End Function

Private Sub Prepare(Optional p_test_no As String = vbNullString)
    
    If Tests Is Nothing Then
        Set Tests = New clsTestAid
        Tests.TestedComp = "mFso"
    End If
    If Not mErH.Regression Then
        mTrc.FileFullName = Tests.TestFolder & "\Test-" & p_test_no & "-ExecTrace.log"
        mTrc.NewFile
    End If

End Sub

Public Sub Test_00_Regression()
' ----------------------------------------------------------------------------
' All results are asserted and there is no intervention required for the whole
' test. When an assertion fails the test procedure will stop and indicates the
' problem with the called procedure.
' ----------------------------------------------------------------------------
    Const PROC = "Test_00_Regression"

    On Error GoTo eh
    Dim sTestStatus As String
    
    '~~ Initialization (must be done prior the first BoP!)
    bRegression = True
    Set Tests = Nothing
    Set Tests = New clsTestAid
    mTrc.FileFullName = Tests.TestFolder & "\Regression.ExecTrace.log"
    mTrc.Title = "Regression Test module ""mVarTrans"""
    mTrc.NewFile
    mErH.Regression = bRegression
    With Tests
        .TestedComp = "mFso"
'        .CleanUp "FailedResult_", "rad" ' cleanup any previous regression test results/remains.
    End With
    mErH.Asserted AppErr(1) ' For the very last test on an error condition
    
    BoP ErrSrc(PROC)
    mFsoTest.Test_01_FileTemp
    mFsoTest.Test_02_Exists
    mFsoTest.Test_04_FilePicked
    mFsoTest.Test_05_FileString
    mFsoTest.Test_06_FileDiffersFromFile
    mFsoTest.Test_08_FileAsArry_ArrayAsFile
    mFsoTest.Test_09_FileSearch
'    mFsoTest.Test_10_FolderIsValidName
'    mFsoTest.Test_11_Folders
'    mFsoTest.Test_12_RenameSubFolders
    
xt: EoP ErrSrc(PROC)
    mTrc.Dsply
    Application.Wait Now() + 0.00001 ' wait to enforce display of the summary in front
    Tests.ResultLogSummary
    mErH.Regression = False
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_01_FileTemp()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_01_FileTemp"

    Dim sTemp As String
    
    Prepare
    BoP ErrSrc(PROC)
    With Tests
        .TestId = "01"
        .Title = "FileTemp service"
        .TestedProc = "FileTemp"
        .TestedProcType = "Function"
        
        .Verification = "Returns a randomly named temporary file"
            .TimerStart
            sTemp = mFso.FileTemp(f_path:=ThisWorkbook.Path)
            .TimerEnd
            .ResultExpected = sTemp
            .Result = sTemp
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_02_Exists()
' ----------------------------------------------------------------------------
' Test of all file exists variants.
' ----------------------------------------------------------------------------
    Const PROC = "Test_02_Exists"
    
    On Error GoTo eh
    Dim cll         As Collection
    Dim sFileName   As String
    
    Prepare "02"
    BoP ErrSrc(PROC)
    
    With Tests
        .TestId = "02"
        .Title = "Exists service"
        .TestedProc = "Exists"
        .TestedProcType = "Function"
        
        .Verification = "Folder not exists"
            .TimerStart
            .Result = mFso.Exists(ThisWorkbook.Path & "x")
            .TimerEnd
            .ResultExpected = False
    
        .Verification = "Folder exists"
            .TimerStart
            .Result = mFso.Exists(ThisWorkbook.Path)
            .TimerEnd
            .ResultExpected = True
    
        .Verification = "File not exists"
            .TimerStart
            .Result = mFso.Exists(e_file:=ThisWorkbook.FullName & "x")
            .TimerEnd
            .ResultExpected = False
    
        .Verification = "File exists"
            .TimerStart
            .Result = mFso.Exists(e_file:=ThisWorkbook.FullName)
            .TimerEnd
            .ResultExpected = True
        
        .Verification = "File by wildcard "
            .TimerStart
            mFso.Exists e_folder:=ThisWorkbook.Path _
                      , e_file:="*.xl*" _
                      , e_result_files:=cll
            .TimerEnd
            .Result = cll.Count
            .ResultExpected = 2
        
        .Verification = "File by wildcard in sub-folders"
            .TimerStart
            mFso.Exists e_folder:=ThisWorkbook.Path & "\CompMan\source" _
                      , e_file:="fMsg.fr*" _
                      , e_result_files:=cll
            .TimerEnd
            .Result = cll.Count
            .ResultExpected = 2
            
        .Verification = "File by wildcard in sub-folders"
            .Result = cll(1).Name
            .ResultExpected = "fMsg.frm"
        
        .Verification = "File by wildcard in sub-folders"
            .Result = cll(2).Name
            .ResultExpected = "fMsg.frx"
        
    End With
                        
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_04_FilePicked()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_04_FilePicked"
    
    On Error GoTo eh
    Dim fl As File

    Prepare
    BoP ErrSrc(PROC)
    
    With Tests
        .TestId = "04"
        .Title = "FilePicked service"
        .TestedProc = "FilePicked"
        .TestedProcType = "Function"
        
        .Verification = "ThisWorkbook is picked"
            .TimerStart
            mFso.FilePicked p_init_path:=ThisWorkbook.Path _
                          , p_filters:="Excel file, *.xl*; All files, *.*" _
                          , p_title:="Test " & .TestId & ": Select the Excel Workbook in this folder (folder preselected by filter)" _
                         , p_file:=fl
            .TimerEnd
            .Result = fl.Path
            .ResultExpected = ThisWorkbook.FullName
        
        .Verification = "No file picked"
            .TimerStart
            mFso.FilePicked p_init_path:=ThisWorkbook.Path _
                          , p_filters:="Excel file, *.xl*; All files, *.*" _
                          , p_title:="Test 2: No file picked (just  t e r m i n a t e  the dialog)" _
                          , p_file:=fl
            .TimerEnd
            .Result = TypeName(fl)
            .ResultExpected = "Nothing"
        
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_05_FileString()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_05_FileString"
    
    On Error GoTo eh
    Dim sFile1  As String
    Dim sTest   As String
    Dim sResult As String
    Dim sSplit  As String
    Dim fle1    As File
    
    Prepare
    BoP ErrSrc(PROC)
    
    With Tests
        .TestId = "05"
        .Title = "String as File, File as String service"
        .TestedProc = "FileAsString, StringAsFile"
        .TestedProcType = "Function"
        
        .Verification = "Write/read one recod"
            sFile1 = mFso.FileTemp()
            sTest = "My string"
            '~~ Write
            mFso.StringAsFile sTest, sFile1
            '~~ Read
            .Result = mFso.FileAsString(sFile1)
            .ResultExpected = sTest
            .TempTestItem = sFile1
        
        .Verification = "Empty file"
            sFile1 = mFso.FileTemp()
            .TempTestItem = sFile1
            sTest = vbNullString
            mFso.StringAsFile sTest, sFile1
            sResult = mFso.FileAsString(sFile1)
            .ResultExpected = vbNullString
    
        .Verification = "Append"
            sFile1 = mFso.FileTemp()
            .TempTestItem = sFile1
            .TimerStart
            mFso.StringAsFile "AAA" & vbCrLf & "BBB", sFile1, False
            mFso.StringAsFile "CCC", sFile1, True
            .TimerEnd
            .Result = mFso.FileAsString(sFile1)
            .ResultExpected = "AAA" & vbCrLf & "BBB" & vbCrLf & "CCC"
        
        .Verification = "Write string to file with append, file as object"
            sFile1 = mFso.FileTemp()
            .TempTestItem = sFile1
            fso.CreateTextFile FileName:=sFile1
            Set fle1 = fso.GetFile(sFile1)
            .TimerStart
            mFso.StringAsFile "AAA" & vbCrLf & "BBB", fle1, False
            mFso.StringAsFile "CCC", fle1, True
            .TimerEnd
            .Result = mFso.FileAsString(fle1)
            .ResultExpected = "AAA" & vbCrLf & "BBB" & vbCrLf & "CCC"
    
    End With
xt: EoP ErrSrc(PROC)
    Tests.CleanUp
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_06_FileDiffersFromFile()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_06_FileDiffersFromFile"
    
    On Error GoTo eh
    Dim f1      As File
    Dim f2      As File
    Dim dctDiff As Dictionary
    Dim sF1     As String
    Dim sF2     As String

    Prepare
    sF1 = mFso.FileTemp
    sF2 = mFso.FileTemp
    BoP ErrSrc(PROC)
    
    With Tests
        .TestId = "06-1"
        .Title = "FileDiffersFromFile service"
        .TestedProc = "FileDiffersFromFile"
        
        .Verification = "Differs = False"
            mFso.StringAsFile "A" & vbCrLf & "B" & vbCrLf & "C", sF1, False
            mFso.StringAsFile "A" & vbCrLf & "B" & vbCrLf & "C", sF2, False
            Set f1 = fso.GetFile(sF1)
            Set f2 = fso.GetFile(sF2)
            .TimerStart
            .Result = mFso.FileDiffersFromFile(f_file_this:=f1 _
                                             , f_file_from:=f2 _
                                             , f_exclude_empty:=True _
                                              )
            .TimerEnd
            .ResultExpected = False
        
        .Verification = "Differs = True"
            .ResultExpected = True
            mFso.StringAsFile "A" & vbCrLf & "B" & vbCrLf & "C", sF1, False
            mFso.StringAsFile "A" & vbCrLf & "B" & vbCrLf & "C" & vbCrLf & "D", sF2, False
            Set f1 = fso.GetFile(sF1)
            Set f2 = fso.GetFile(sF2)
            .TimerStart
            .Result = mFso.FileDiffersFromFile(f_file_this:=f1 _
                                             , f_file_from:=f2 _
                                             , f_exclude_empty:=True _
                                              )
            .TimerEnd
        
        .Verification = "Differs = True"
            .ResultExpected = True
            ' Test 3: Differs.Count = 1
            mFso.StringAsFile "A" & vbCrLf & "B" & vbCrLf & "C" & vbCrLf & "D", sF1, False
            mFso.StringAsFile "A" & vbCrLf & "B" & vbCrLf & "C", sF2, False
            Set f1 = fso.GetFile(sF1)
            Set f2 = fso.GetFile(sF2)
            .TimerStart
            .Result = mFso.FileDiffersFromFile(f_file_this:=f1 _
                                             , f_file_from:=f2 _
                                             , f_exclude_empty:=True _
                                              )
            .TimerEnd
        
        .Verification = "Differs = True"
            mFso.StringAsFile "A" & vbCrLf & "B" & vbCrLf & "C", sF1, False
            mFso.StringAsFile "A" & vbCrLf & "X" & vbCrLf & "C", sF2, False
            Set f1 = fso.GetFile(sF1)
            Set f2 = fso.GetFile(sF2)
            .TimerStart
            .Result = mFso.FileDiffersFromFile(f_file_this:=f1 _
                                             , f_file_from:=f2 _
                                             , f_exclude_empty:=True _
                                              )
            .TimerEnd
            .ResultExpected = True
    
    End With
    
xt: EoP ErrSrc(PROC)
    If fso.FileExists(sF1) Then fso.DeleteFile (sF1)
    If fso.FileExists(sF2) Then fso.DeleteFile (sF2)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub
  
Public Sub Test_08_FileAsArry_ArrayAsFile()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_08_FileArry_Get_Let"
    
    On Error GoTo eh
    Dim arry    As Variant
    Dim fle     As File
    Dim sFile   As String
    
    Prepare
    BoP ErrSrc(PROC)
    With Tests
        .CleanUp
        .TestId = "08"
        .Title = "File as array, array as file service"
        .TestedProcType = "Function"
        
        .Verification = "Write file from array"
            .TestedProc = "ArryAsFile"
            sFile = .TempFileFullName
            .TimerStart
            Set fle = mFso.ArrayAsFile(Array("xxx", vbNullString, "yyy"), sFile)
            .TimerEnd
            .Result = mFso.FileAsArray(fle)
            .ResultExpected = Array("xxx", vbNullString, "yyy")
         
        .Verification = "Read file to array exclude empty false"
            .TestedProc = "FileAsArry"
            sFile = .TempFileFullName
            Set fle = mFso.ArrayAsFile(Array("xxx", vbNullString, "yyy"), sFile)
            .TimerStart
            arry = mFso.FileAsArray(fle, False)
            .TimerEnd
            .Result = UBound(arry) - LBound(arry) + 1
            .ResultExpected = 3
            
        .Verification = "Retrurn a file as array with empty items excluded"
            .TestedProc = "ArryAsFile"
            sFile = .TempFileFullName
            Set fle = mFso.ArrayAsFile(Array("xxx", vbNullString, , "yyy"), sFile)
            arry = mFso.FileAsArray(fle, True)
            .Result = UBound(arry) - LBound(arry) + 1
            .ResultExpected = 2
            
            .CleanUp
    End With
        
xt: EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_09_FileSearch()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_09_FileSearch"
    
    On Error GoTo eh
    Dim cll As Collection
    
    BoP ErrSrc(PROC)
    Prepare
    
    With Tests
        .TestId = "09"
        .Title = "Test of FileSearch"
        .TestedProc = "FileSearch"
        .TestedProcType = "Function"
        
        .Verification = "File search with wildcards, including subfolders, several files found"
            Set cll = mFso.FilesSearch(f_root:=fso.GetFolder(ThisWorkbook.Path).ParentFolder.Path & "\Common-Components\" _
                                    , f_mask:="*.bas*" _
                                    , f_in_subfolders:=True _
                                      )
            .Result = cll.Count
            .ResultExpected = cll.Count
    
        .Verification = "File search with wildcards, subfolders excluded, no files found"
            .Result = mFso.FilesSearch(f_root:="e:\Ablage\Excel VBA\DevAndTest\Common" _
                                     , f_mask:="*CompMan*.frx" _
                                     , f_in_subfolders:=False _
                                      ).Count
            .ResultExpected = 0
            
    End With
    
xt: EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_10_FolderIsValidName()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_10_FolderIsValidName"
    
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

Public Sub Test_11_Folders()
    Const PROC = "Test_11_Folders"
    
    Dim TestFolder As String
    
    BoP ErrSrc(PROC)
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
    EoP ErrSrc(PROC)
        
End Sub

Public Sub Test_12_RenameSubFolders()
    Const PROC = "Test_12_RenameSubFolders"
    
    Dim cllRenamed      As Collection
    Dim sFolderOldPath  As String
    Dim sFolderNewPath  As String
    Dim sFolderNewName  As String
    Dim sFolderOldName  As String
    Dim sFolderRootPath     As String
    
    BoP ErrSrc(PROC)
    sFolderRootPath = ThisWorkbook.Path & "\Test"
    sFolderOldName = "SubFolder"
    sFolderNewName = "SubFolder_renamed"
    
    '~~ Test 1: Rename one sub-folder only
    Set cllRenamed = New Collection
    Test_12_RenameSubFolders_Prepare sFolderRootPath, sFolderOldName
    
    mFso.RenameSubFolders sFolderRootPath & "\Test1", sFolderOldName, sFolderNewName, cllRenamed
    Debug.Assert cllRenamed.Count = 1
    Debug.Assert cllRenamed(1).Path = sFolderRootPath & "\Test1\SubFolder_renamed"

    '~~ Test 2: Rename all (2) sub-folders
    Set cllRenamed = New Collection
    Test_12_RenameSubFolders_Prepare sFolderRootPath, sFolderOldName
    sFolderRootPath = ThisWorkbook.Path & "\Test"
    
    mFso.RenameSubFolders sFolderRootPath, sFolderOldName, sFolderNewName, cllRenamed
    Debug.Assert cllRenamed.Count = 2
    Debug.Assert cllRenamed(1).Path = sFolderRootPath & "\Test1\SubFolder_renamed"
    Debug.Assert cllRenamed(2).Path = sFolderRootPath & "\Test2\SubFolder_renamed"
    
    Set cllRenamed = Nothing
    EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_12_RenameSubFolders_Prepare(ByVal s_path As String, _
                                             ByVal s_folder_old_name As String)
                                             
    With fso
        If .FolderExists(s_path) Then .DeleteFolder s_path
        .CreateFolder s_path
        .CreateFolder s_path & "\Test1"
        .CreateFolder s_path & "\Test2"
        .CreateFolder s_path & "\Test1\" & s_folder_old_name
        .CreateFolder s_path & "\Test2\" & s_folder_old_name
    End With
    
End Sub

