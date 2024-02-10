Attribute VB_Name = "mFsoTest"
Option Explicit
' ----------------------------------------------------------------
' Standard Module mFsoTest: Test of all services of the module.
'
' ----------------------------------------------------------------
Private Const SECTION_NAME = "Section-" ' for PrivateProfile services test
Private Const VALUE_NAME = "-Name-"     ' for PrivateProfile services test
Private Const VALUE_STRING = "-Value-"  ' for PrivateProfile services test
    
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
    ErrSrc = "mFsoTest." & sProc
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
    mTrc.FileName = "RegressionTest.ExecTrace.log"
    mTrc.Title = "Regression Test module mFso"
    mTrc.NewFile
    mErH.Regression = True
    
    mBasic.BoP ErrSrc(PROC)
    sTestStatus = "mFso Regression Test: "

    mErH.Asserted AppErr(1) ' For the very last test on an error condition
    mFsoTest.Test_00_Regression_Common_Services
    mFsoTest.Test_00_Regression_PrivateProfile_Services

xt: TestProc_RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    mErH.Regression = False
    mTrc.Dsply
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
    mBasic.BoP ErrSrc(PROC)
    
    mErH.Asserted AppErr(1) ' For the very last test on an error condition
    mFsoTest.Test_01_FileTemp
    mFsoTest.Test_02_File_Exists
    mFsoTest.Test_03_FilePathSplit
    mFsoTest.Test_07_FilePicked
    mFsoTest.Test_08_FileString_Let_Get
    mFsoTest.Test_09_FileDiffersFromFile_False
    mFsoTest.Test_09_FileDiffersFromFile
    mFsoTest.Test_10_FileArry_Get_Let
    mFsoTest.Test_11_FileSearch
    mFsoTest.Test_12_FolderIsValidName
    mFsoTest.Test_13_Folders
    mFsoTest.Test_14_RenameSubFolders

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
    
    mBasic.BoP ErrSrc(PROC)
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
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
    
    mBasic.BoP ErrSrc(PROC)
    sTemp = mFso.FileTemp(f_path:=ThisWorkbook.Path)
    sTemp = mFso.FileTemp()
    mBasic.EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_02_File_Exists()
' ----------------------------------------------------------------------------
' Test of all file exists variants.
' ----------------------------------------------------------------------------
    Const PROC = "Test_02_File_Exists"
    
    On Error GoTo eh
    Dim cll         As Collection
    Dim sFileName   As String
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Folder exists
    Debug.Assert mFso.Exists(x_folder:=ThisWorkbook.Path & "x") = False
    Debug.Assert mFso.Exists(x_folder:=ThisWorkbook.Path) = True
    
    '~~ File exists
    Debug.Assert mFso.Exists(x_file:=ThisWorkbook.FullName & "x") = False
    Debug.Assert mFso.Exists(x_file:=ThisWorkbook.FullName) = True

    '~~ Section exists
    sFileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues)
    Debug.Assert Not mFso.Exists(x_file:=sFileName _
                               , x_section:=TestProc_SectionName(2) & "x" _
                                )
    Debug.Assert mFso.Exists(x_file:=sFileName _
                           , x_section:=TestProc_SectionName(2) _
                            )
    '~~ Value-Name exists
    Debug.Assert Not mFso.Exists(x_file:=sFileName _
                               , x_section:=TestProc_SectionName(2) _
                               , x_value_name:=TestProc_ValueName(2, 2) & "x" _
                                )
    Debug.Assert mFso.Exists(x_file:=sFileName _
                            , x_section:=TestProc_SectionName(2) _
                            , x_value_name:=TestProc_ValueName(2, 2) _
                             )

    '~~ File by wildcard, in any sub-folder, exactly one
    Debug.Assert mFso.Exists(x_folder:=ThisWorkbook.Path _
                           , x_file:="*.xl*" _
                           , x_result_files:=cll)
    Debug.Assert cll.Count = 1
    Debug.Assert cll(1).Path = ThisWorkbook.FullName
            
    '~~ File by wildcard, in any sub-folder, more than one
    Debug.Assert mFso.Exists(x_folder:=ThisWorkbook.Path _
                           , x_file:="fMsg.fr*" _
                           , x_result_files:=cll) = True
    Debug.Assert cll.Count = 2
    Debug.Assert cll(1).Name = "fMsg.frm"
    Debug.Assert cll(2).Name = "fMsg.frx"
                        
xt: TestProc_RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
    
    mBasic.BoP ErrSrc(PROC)
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
    Debug.Assert f = vbNullString ' any last element with a . is interpreted as file name
    
    p = ThisWorkbook.Path & "x"
    mFso.FilePathSplit p, f
    Debug.Assert p = vbNullString
    Debug.Assert f = vbNullString

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_07_FilePicked()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_07_FilePicked"
    
    On Error GoTo eh
    Dim fl As File

    mBasic.BoP ErrSrc(PROC)
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
    
xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_08_FileString_Let_Get()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_08_FileString_Let_Get"
    
    On Error GoTo eh
    Dim sFl     As String
    Dim sTest   As String
    Dim sResult As String
    Dim sSplit  As String
    Dim oFl     As File
    
    Test_Status = ErrSrc(PROC)
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Test 1: Write one recod
    sFl = mFso.FileTemp()
    sTest = "My string"
    mFso.FileString(f_file_full_name:=sFl _
                  , f_append:=False _
                   ) = sTest
    sResult = mFso.FileString(f_file_full_name:=sFl, f_split:=sSplit)
    Debug.Assert Split(sResult, sSplit)(0) = sTest
    FSo.DeleteFile sFl
    
    '~~ Test 2: Empty file
    sFl = mFso.FileTemp()
    sTest = vbNullString
    mFso.FileString(f_file_full_name:=sFl, f_append:=False) = sTest
    sResult = mFso.FileString(f_file_full_name:=sFl, f_split:=sSplit)
    Debug.Assert sResult = vbNullString
    FSo.DeleteFile sFl

    '~~ Test 3: Append
    sFl = mFso.FileTemp()
    mFso.FileString(f_file_full_name:=sFl, f_append:=False) = "AAA" & vbCrLf & "BBB"
    mFso.FileString(f_file_full_name:=sFl, f_append:=True) = "CCC"
    sResult = mFso.FileString(f_file_full_name:=sFl, f_split:=sSplit)
    Debug.Assert Split(sResult, sSplit)(0) = "AAA"
    Debug.Assert Split(sResult, sSplit)(1) = "BBB"
    Debug.Assert Split(sResult, sSplit)(2) = "CCC"
    FSo.DeleteFile sFl

    '~~ Test 4: Write with append and read with file as object
    sFl = mFso.FileTemp()
    FSo.CreateTextFile FileName:=sFl
    Set oFl = FSo.GetFile(sFl)
    sFl = oFl.Path
    mFso.FileString(f_file_full_name:=oFl, f_append:=False) = "AAA" & vbCrLf & "BBB"
    mFso.FileString(f_file_full_name:=oFl, f_append:=True) = "CCC"
    sResult = mFso.FileString(f_file_full_name:=oFl, f_split:=sSplit)
    Debug.Assert Split(sResult, sSplit)(0) = "AAA"
    Debug.Assert Split(sResult, sSplit)(1) = "BBB"
    Debug.Assert Split(sResult, sSplit)(2) = "CCC"
    FSo.DeleteFile sFl

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_09_FileDiffersFromFile()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_09_FileDiffersFromFile"
    
    On Error GoTo eh
    Dim f1      As File
    Dim f2      As File
    Dim dctDiff As Dictionary
    Dim sF1     As String
    Dim sF2     As String

    mBasic.BoP ErrSrc(PROC)
    Test_Status = ErrSrc(PROC)
    
    sF1 = mFso.FileTemp
    sF2 = mFso.FileTemp

    mBasic.BoP ErrSrc(PROC)
    ' Prepare
    mFso.FileString(f_file_full_name:=sF1, f_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    mFso.FileString(f_file_full_name:=sF2, f_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    Set f1 = FSo.GetFile(sF1)
    Set f2 = FSo.GetFile(sF2)

    ' Test 1: Differs = False
    Debug.Assert Not mFso.FileDiffersFromFile(f_file_this:=f1 _
                                            , f_file_from:=f2 _
                                            , f_exclude_empty:=True _
                                             )
    

    ' Test 2: Differs.Count = 1
    mFso.FileString(f_file_full_name:=sF1, f_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    mFso.FileString(f_file_full_name:=sF2, f_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C" & vbCrLf & "D"
    Set f1 = FSo.GetFile(sF1)
    Set f2 = FSo.GetFile(sF2)
    
    Debug.Assert mFso.FileDiffersFromFile(f_file_this:=f1 _
                                        , f_file_from:=f2 _
                                        , f_exclude_empty:=True _
                                          )
    
    ' Test 3: Differs.Count = 1
    mFso.FileString(f_file_full_name:=sF1, f_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C" & vbCrLf & "D"
    mFso.FileString(f_file_full_name:=sF2, f_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    Set f1 = FSo.GetFile(sF1)
    Set f2 = FSo.GetFile(sF2)
    
    Debug.Assert mFso.FileDiffersFromFile(f_file_this:=f1 _
                                        , f_file_from:=f2 _
                                        , f_exclude_empty:=True _
                                         )
    
    ' Test 4: Differs.Count = 1
    mFso.FileString(f_file_full_name:=sF1, f_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    mFso.FileString(f_file_full_name:=sF2, f_append:=False) = "A" & vbCrLf & "X" & vbCrLf & "C"
    Set f1 = FSo.GetFile(sF1)
    Set f2 = FSo.GetFile(sF2)
    
    Debug.Assert mFso.FileDiffersFromFile(f_file_this:=f1 _
                                        , f_file_from:=f2 _
                                        , f_exclude_empty:=True _
                                         )
xt: mBasic.EoP ErrSrc(PROC)
    If FSo.FileExists(sF1) Then FSo.DeleteFile (sF1)
    If FSo.FileExists(sF2) Then FSo.DeleteFile (sF2)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_09_FileDiffersFromFile_False()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_09_FileDiffersFromFile_False"
    
    On Error GoTo eh
    Dim sFileName   As String
    Dim f1          As File
    Dim f2          As File
    Dim dctDiff     As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    Test_Status = ErrSrc(PROC)
    ' Prepare
    sFileName = FSo.GetFolder(ThisWorkbook.Path).ParentFolder.Path & "\Common-Components\mFso.bas"
    Set f1 = FSo.GetFile(sFileName)
    Set f2 = FSo.GetFile(sFileName)
    
    ' Test
    Debug.Assert Not mFso.FileDiffersFromFile(f_file_this:=f1 _
                                            , f_file_from:=f2 _
                                            , f_exclude_empty:=True)

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

  
Public Sub Test_10_FileArry_Get_Let()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_10_FileArry_Get_Let"
    
    On Error GoTo eh
    Dim sFile1      As String
    Dim sFile2      As String
    Dim lInclEmpty  As Long
    Dim lEmpty1     As Long
    Dim lExclEmpty  As Long
    Dim lEmpty2     As Long
    Dim a           As Variant
    Dim v           As Variant
    
    mBasic.BoP ErrSrc(PROC)
    Test_Status = ErrSrc(PROC)
    
    sFile1 = FSo.GetFolder(ThisWorkbook.Path).ParentFolder.Path & "Common-Components\mFso.bas"
    sFile2 = FSo.GetFolder(ThisWorkbook.Path).ParentFolder.Path & "Common-Components\mFso.bas"
    
    sFile1 = mFso.FileTemp()
    sFile2 = mFso.FileTemp()
    
    '~~ Write to lines to sFile1
    mFso.FileString(sFile1) = "xxx" & vbCrLf & "" & "yyy"
    
    '~~ Get the two lines as Array
    a = mFso.FileArry(f_file_full_name:=sFile1)
    Debug.Assert a(LBound(a)) = "xxx"
    Debug.Assert a(UBound(a)) = "yyy"

    '~~ Write array to file-2
    mFso.FileArry(f_file_full_name:=sFile2 _
             , f_split:=vbCrLf _
              ) = a
    Debug.Assert mFso.FileDiffersFromFile(f_file_this:=sFile1 _
                                        , f_file_from:=sFile2) = False

    '~~ Count empty records when array contains all text lines
    a = mFso.FileArry(f_file_full_name:=sFile1, f_exclude_empty:=False)
    lInclEmpty = UBound(a) + 1
    lEmpty1 = 0
    For Each v In a
        If VBA.Trim$(v) = vbNullString Then lEmpty1 = lEmpty1 + 1
        If VBA.Len(Trim$(v)) = 0 Then lEmpty2 = lEmpty2 + 1
    Next v
    
    '~~ Count empty records
    a = mFso.FileArry(f_file_full_name:=sFile1, f_exclude_empty:=True)
    lExclEmpty = UBound(a) + 1
    Debug.Assert lExclEmpty = lInclEmpty - lEmpty1
    
xt: With FSo
        .DeleteFile sFile1
        If .FileExists(sFile2) Then .DeleteFile sFile2
    End With
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_11_FileSearch()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_11_FileSearch"
    
    On Error GoTo eh
    Dim cll As Collection
    
    mBasic.BoP ErrSrc(PROC)
    Test_Status = ErrSrc(PROC)
    
    '~~ Test 1: Including subfolders, several files found
    Set cll = mFso.FilesSearch(f_root:=FSo.GetFolder(ThisWorkbook.Path).ParentFolder.Path & "\Common-Components\" _
                             , f_mask:="*.bas*" _
                             , f_stop_after:=5 _
                              )
    Debug.Assert cll.Count > 2

    '~~ Test 2: Not including subfolders, no files found
    Set cll = mFso.FilesSearch(f_root:="e:\Ablage\Excel VBA\DevAndTest\Common" _
                             , f_mask:="*CompMan*.frx" _
                             , f_stop_after:=5 _
                             , f_in_subfolders:=False _
                              )
    Debug.Assert cll.Count = 0

xt: mBasic.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_12_FolderIsValidName()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_12_FolderIsValidName"
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Test 1: Valid Folder Name
    Debug.Assert mFso.FolderIsValidName(ThisWorkbook.Path) = True        ' a valid folder is a valid file name as well
    Debug.Assert mFso.FolderIsValidName(ThisWorkbook.FullName) = True
    Debug.Assert mFso.FolderIsValidName(ThisWorkbook.Name) = False
    Debug.Assert mFso.FolderIsValidName("c:\LP?1") = False

    '~~ Test 2: Valid File Name
    Debug.Assert mFso.FileIsValidName(ThisWorkbook.Name) = True
    Debug.Assert mFso.FileIsValidName(ThisWorkbook.Name & "?") = False
    mBasic.EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_13_Folders()
    Const PROC = "Test_13_Folders"
    
    Dim TestFolder As String
    
    mBasic.BoP ErrSrc(PROC)
    TestFolder = FSo.GetFolder(ThisWorkbook.Path).ParentFolder.Path
    
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
    mBasic.EoP ErrSrc(PROC)
        
End Sub

Public Sub Test_14_RenameSubFolders()
    Const PROC = "Test_14_RenameSubFolders"
    
    Dim cllRenamed      As Collection
    Dim sFolderOldPath  As String
    Dim sFolderNewPath  As String
    Dim sFolderNewName  As String
    Dim sFolderOldName  As String
    Dim sFolderRootPath     As String
    
    mBasic.BoP ErrSrc(PROC)
    sFolderRootPath = ThisWorkbook.Path & "\Test"
    sFolderOldName = "SubFolder"
    sFolderNewName = "SubFolder_renamed"
    
    '~~ Test 1: Rename one sub-folder only
    Set cllRenamed = New Collection
    Test_14_RenameSubFolders_Prepare sFolderRootPath, sFolderOldName
    
    mFso.RenameSubFolders sFolderRootPath & "\Test1", sFolderOldName, sFolderNewName, cllRenamed
    Debug.Assert cllRenamed.Count = 1
    Debug.Assert cllRenamed(1).Path = sFolderRootPath & "\Test1\SubFolder_renamed"

    '~~ Test 2: Rename all (2) sub-folders
    Set cllRenamed = New Collection
    Test_14_RenameSubFolders_Prepare sFolderRootPath, sFolderOldName
    sFolderRootPath = ThisWorkbook.Path & "\Test"
    
    mFso.RenameSubFolders sFolderRootPath, sFolderOldName, sFolderNewName, cllRenamed
    Debug.Assert cllRenamed.Count = 2
    Debug.Assert cllRenamed(1).Path = sFolderRootPath & "\Test1\SubFolder_renamed"
    Debug.Assert cllRenamed(2).Path = sFolderRootPath & "\Test2\SubFolder_renamed"
    
    Set cllRenamed = Nothing
    mBasic.EoP ErrSrc(PROC)
    
End Sub

Private Sub Test_14_RenameSubFolders_Prepare(ByVal s_path As String, _
                                             ByVal s_folder_old_name As String)
                                             
    With FSo
        If .FolderExists(s_path) Then .DeleteFolder s_path
        .CreateFolder s_path
        .CreateFolder s_path & "\Test1"
        .CreateFolder s_path & "\Test2"
        .CreateFolder s_path & "\Test1\" & s_folder_old_name
        .CreateFolder s_path & "\Test2\" & s_folder_old_name
    End With
    
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
    sFileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues)
    
    Set dct = mFso.PPsectionNames(sFileName)
    Debug.Assert dct.Count = 3
    Debug.Assert dct.Items()(0) = TestProc_SectionName(1)
    Debug.Assert dct.Items()(1) = TestProc_SectionName(2)
    Debug.Assert dct.Items()(2) = TestProc_SectionName(3)

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
    sFileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues)
    
    Set dct = mFso.PPvalues(pp_file:=sFileName, pp_section:=TestProc_SectionName(2))
    mBasic.EoP ErrSrc(PROC)
    Debug.Assert dct.Count = 3
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
    Dim sFileName   As String
    Dim cyValue     As Currency: cyValue = 12345.6789
    Dim cyResult    As Currency
    
    mBasic.BoP ErrSrc(PROC)
    '~~ Test preparation
    sFileName = TestProc_TempFile
            
    '~~ Test 1: Read non-existing value from a non-existing file
    Debug.Assert mFso.PPvalue(pp_file:=sFileName _
                           , pp_section:="Any" _
                           , pp_value_name:="Any" _
                            ) = vbNullString
    
    '~~ Test 2: Write values
    mFso.PPvalue(pp_file:=sFileName, pp_section:=TestProc_SectionName(1), pp_value_name:=TestProc_ValueName(1, 1)) = TestProc_ValueString(1, 1)
    mFso.PPvalue(pp_file:=sFileName, pp_section:=TestProc_SectionName(1), pp_value_name:=TestProc_ValueName(1, 2)) = TestProc_ValueString(1, 2)
    mFso.PPvalue(pp_file:=sFileName, pp_section:=TestProc_SectionName(2), pp_value_name:=TestProc_ValueName(2, 1)) = TestProc_ValueString(2, 1)
    mFso.PPvalue(pp_file:=sFileName, pp_section:=TestProc_SectionName(2), pp_value_name:=TestProc_ValueName(2, 2)) = cyValue
    
    '~~ Test 2: Assert written values
    Debug.Assert mFso.PPvalue(pp_file:=sFileName, pp_section:=TestProc_SectionName(1), pp_value_name:=TestProc_ValueName(1, 1)) = TestProc_ValueString(1, 1)
    Debug.Assert mFso.PPvalue(pp_file:=sFileName, pp_section:=TestProc_SectionName(1), pp_value_name:=TestProc_ValueName(1, 2)) = TestProc_ValueString(1, 2)
    Debug.Assert mFso.PPvalue(pp_file:=sFileName, pp_section:=TestProc_SectionName(2), pp_value_name:=TestProc_ValueName(2, 1)) = TestProc_ValueString(2, 1)
    cyResult = mFso.PPvalue(pp_file:=sFileName, pp_section:=TestProc_SectionName(2), pp_value_name:=TestProc_ValueName(2, 2))
    Debug.Assert cyResult = cyValue
    Debug.Assert VarType(cyResult) = vbCurrency
    
xt: TestProc_RemoveTestFiles
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
    sFileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues)
    
    '~~ Test 1: All values of one section
    Set dct = mFso.PPvalues(pp_file:=sFileName _
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
    Debug.Assert mFso.PPvalues(sFileName, vbNullString).Count = 0

    '~~ Test 3: Section does not exist
    Debug.Assert mFso.PPvalues(sFileName, "xxxxxxx").Count = 0

xt: TestProc_RemoveTestFiles
    Set dct = Nothing
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
    sFileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues)
       
    '~~ Section not exists
    Debug.Assert mFso.PPsectionExists(pp_file:=sFileName _
                                              , pp_section:=TestProc_SectionName(100) _
                                               ) = False
    '~~ Section exists
    Debug.Assert mFso.PPsectionExists(pp_file:=sFileName _
                                    , pp_section:=TestProc_SectionName(9) _
                                     ) = True
    '~~ Value-Name exists
    Debug.Assert mFso.PPvalueExists(pp_file:=sFileName _
                                  , pp_section:=TestProc_SectionName(7) _
                                  , pp_value_name:=TestProc_ValueName(7, 3) _
                                   ) = True
    '~~ Value-Name not exists
    Debug.Assert mFso.PPvalueExists(pp_file:=sFileName _
                                  , pp_section:=TestProc_SectionName(7) _
                                  , pp_value_name:=TestProc_ValueName(6, 3) _
                                   ) = False
    
xt: TestProc_RemoveTestFiles
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
' - Test_91_PrivateProfile_SectionsNames (mFso.PPsectionNames)
' Iplicitely tested are:
' - mFso.PPsections Get and Let
' ----------------------------------------------------------------------------
    Const PROC = "Test_97_PrivateProfile_SectionsCopy"
    
    On Error GoTo eh
    Dim SourceFile      As String
    Dim TargetFile      As String
    Dim sSectionName    As String
    Dim dct             As Dictionary
    
    mBasic.BoP ErrSrc(PROC)
    
    '~~ Test 1a ------------------------------------
    '~~ Copy a specific section to a new target file
    SourceFile = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues) ' prepare PrivateProfile test file
    TargetFile = mFso.FileTemp(f_extension:=".dat")
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
    FSo.DeleteFile SourceFile
    FSo.DeleteFile TargetFile
    
    '~~ Test 3 -------------------------------
    '~~ Copy all sections to a new target file (will be re-ordered ascending thereby)
    SourceFile = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues) ' prepare PrivateProfile test file
    TargetFile = mFso.FileTemp(f_extension:=".dat")
    mFso.PPsectionsCopy pp_source:=SourceFile _
                                , pp_target:=TargetFile _
                                , pp_sections:=mFso.PPsectionNames(SourceFile) _
                                , pp_merge:=False
    '~~ Assert result
    Debug.Assert mFso.FileArry(TargetFile)(0) = "[" & TestProc_SectionName(1) & "]"
    FSo.DeleteFile SourceFile
    FSo.DeleteFile TargetFile
            
xt: TestProc_RemoveTestFiles
    mBasic.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mBasic.ErrMsg(ErrSrc(PROC))
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
    Dim sFileName   As String
    Dim vFile       As Variant
    
    mBasic.BoP ErrSrc(PROC)
    
    sFileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues) ' prepare PrivateProfile test file
    vFile = mFso.FileArry(sFileName)
    Debug.Assert vFile(0) = "[" & TestProc_SectionName(10) & "]"
    Debug.Assert vFile(1) = TestProc_ValueName(10, 10) & "=" & TestProc_ValueString(10, 10)
    
    mFso.PPreorg sFileName
    '~~ Assert result
    vFile = mFso.FileArry(sFileName)
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

Public Sub Test_99_PrivateProfile_PPreorg_WithNoFileProvided()
' ----------------------------------------------------------------------------
'
' ----------------------------------------------------------------------------
    Const PROC = "Test_99_PrivateProfile_NoFileProvided"
    
    On Error GoTo eh
    Dim sFileName   As String
    Dim vFile       As Variant
    
    mBasic.BoP ErrSrc(PROC)
    sFileName = TestProc_PrivateProfile_File(mTest.lTestSections, mTest.lTestValues) ' prepare PrivateProfile test file
    
    mFso.PPreorg sFileName
    '~~ Assert result
    vFile = mFso.FileArry(sFileName)
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

