VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPrivProfFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Option Compare Text
' ----------------------------------------------------------------------------
' Class Module clsPrivProfFile: Services regarding Private Profile Files. The
' ============================= componet works autonomous and does not require
' any other components.
'
' Public properties and services:
' -------------------------------
' Exists          Universal existence function. Returns TRUE when the folde
'                 folder, file, section ([xxxx]), or value name exists.
' Reorg           Reorganizes all sections and their value-names in
'                 ascending sequence.
' RemoveNames     Removes provided value names, in a given Private
'                 Properties File, when provided in a specific section,
'                 else in all sections.
' SectionExists   Returns TRUE when a given section exists in a given
'                 Private Profile File.
' SectionNames    Returns a Dictionary of all section names [...] in a
'                 Private Profile File whereby each item is a Dictionary of
'                 value-names as key and the value as item, and the key of
'                 the item is the section name.
' Sections    r/w - Returns a Dictionary of all section names [...] in a
'                   Private Profile File whereby each item is a Dictionary
'                   of value-names as key and the value as item, and the key
'                   of the item is the section name
'                 - Writes a Private Profile File with the content of a
'                   Dictionary.
' RemoveSections  Removes the sections provided via by their given name.
'                 Sections not existing are ignored.
' Value       r/w - Reads from a Private Profile File a value with a
'                   provided value-name from a provided section
'                 - Writes to a Private Profile File a value with a provided
'                   value-name into a provided section
' ValueNameExists Returns TRUE when a value-name exists in a provided
'                 Private Profile File's section.
' ValueNameRename Function replaces an old value name with a new one
'                 either in a specific section or in all sections when no
'                 specific section is provided. Optionally not reorgs the
'                 file, returns True when at least one name has been
'                 replaced.
' ValueNames      Returns a Dictionary of all value-names within given
'                 sections in a Private Profile File with the value-name and
'                 the section name as key (<name>[section]) and the value as
'                 item, the names in ascending order in a Dictionary.
'                 Section names may be provided as a comma delimited string,
'                 a Dictionary or Collection. Non existing sections are
'                 ignored.
' Values          Returns the value-names and values of a given section in a
'                 Private Profile File as Dictionary with the value-name as
'                 the key (in ascending order) and the value as item.
'
' Requires References to: Microsoft Scripting Runtine
'                         Windows Script Host Object Model
'
' Uses no other components. Will use optionally mErH, fMsg/mMsg when installed
' and activated (Cond. Comp. Args. `mErH = 1 : mMsg = 1`).
'
' W. Rauschenberger, Berlin Jan 2024
' See also https://github.com/warbe-maker/VBA-File-System-Objects.
' ----------------------------------------------------------------------------
Public FSo                      As FileSystemObject

Private Const GITHUB_REPO_URL   As String = "https://github.com/warbe-maker/VBA-File-System-Objects"
Private Const WIN_NORMAL = 1         'Open Normal
Private Const ERROR_SUCCESS = 32&
Private Const ERROR_NO_ASSOC = 31&
Private Const ERROR_OUT_OF_MEM = 0&
Private Const ERROR_FILE_NOT_FOUND = 2&
Private Const ERROR_PATH_NOT_FOUND = 3&
Private Const ERROR_BAD_FORMAT = 11&

#If mMsg = 0 Then
    ' ------------------------------------------------------------------------
    ' The 'minimum error handling' approach implemented with this module and
    ' provided by the ErrMsg function uses the VBA.MsgBox to display an error
    ' message which includes a debugging option to resume the error line
    ' provided the Conditional Compile Argument 'Debugging = 1'.
    ' This declaration allows the mTrc module to work completely autonomous.
    ' It becomes obsolete when the mMsg/fMsg module is installed 1) which must
    ' be indicated by the Conditional Compile Argument mMsg = 1
    '
    ' 1) See https://github.com/warbe-maker/Common-VBA-Message-Service for
    '    how to install an use.
    ' ------------------------------------------------------------------------
    Private Const vbResumeOk As Long = 7 ' Buttons value in mMsg.ErrMsg (pass on not supported)
    Private Const vbResume   As Long = 6 ' return value (equates to vbYes)
#End If

Private Declare PtrSafe Function WritePrivateProfileString _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal lpw_ApplicationName As String, _
                ByVal lpw_KeyName As String, _
                ByVal lpw_String As String, _
                ByVal lpw_FileName As String) As Long
                
Private Declare PtrSafe Function GetPrivateProfileString _
                Lib "kernel32" Alias "GetPrivateProfileStringA" _
               (ByVal lpg_ApplicationName As String, _
                ByVal lpg_KeyName As String, _
                ByVal lpg_Default As String, _
                ByVal lpg_ReturnedString As String, _
                ByVal nSize As Long, _
                ByVal lpg_FileName As String) As Long

Private Declare PtrSafe Function DeletePrivateProfileSection _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal Section As String, _
                ByVal NoKey As Long, _
                ByVal NoSetting As Long, _
                ByVal Name As String) As Long

Private Declare PtrSafe Function DeletePrivateProfileKey _
                Lib "kernel32" Alias "WritePrivateProfileStringA" _
               (ByVal Section As String, _
                ByVal Key As String, _
                ByVal Setting As Long, _
                ByVal Name As String) As Long
                 
Private Declare PtrSafe Function GetPrivateProfileSectionNames _
                Lib "kernel32.dll" Alias "GetPrivateProfileSectionNamesA" _
               (ByVal lpszReturnBuffer As String, _
                ByVal nSize As Long, _
                ByVal lpName As String) As Long

Private Declare PtrSafe Function apiShellExecute _
                Lib "shell32.dll" Alias "ShellExecuteA" _
               (ByVal hWnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Private sFileName   As String
Private sSection    As String
Private dctValues   As Dictionary
Private dctSections As Dictionary
Private Reg         As RegExp
Private dctFile     As Dictionary

Private Sub Class_Initialize()

    Set Reg = New RegExp
    Set FSo = New FileSystemObject
    Set dctFile = New Dictionary
    '~~ Initialize the default Private Profile File name
    sFileName = ThisWorkbook.Path & "\" & FSo.GetBaseName(ThisWorkbook.Name) & ".dat"
       
End Sub

Private Sub FileAsDict(ByVal f_file As String)
' ----------------------------------------------------------------------------
' Adds the current file (sFileName) as Dictionary to the dctFile Dictionary
' when yet not added.
' ----------------------------------------------------------------------------

    Dim arr         As Variant
    Dim dctValues   As Dictionary
    Dim dct         As Dictionary
    Dim i           As Long
    Dim s           As String
    Dim sName       As String
    Dim sValue      As String
    
    If dctFile.Exists(sFileName) Then Exit Sub
    
    arr = Split(FileString(sFileName), vbCrLf)
    Set dct = New Dictionary
    
    For i = LBound(arr) To UBound(arr)
        If ItemIsSection(arr(i), s) Then
            If Not dctValues Is Nothing Then
                dct.Add sSection, dctValues
                Set dctValues = Nothing
            End If
            sSection = s
            Set dctValues = New Dictionary
        ElseIf ItemIsValue(arr(i), sName, sValue) Then
            dctValues.Add sName, sValue
        End If
    Next i
    If Not dctValues Is Nothing Then
        dct.Add sSection, dctValues
        Set dctValues = Nothing
    End If
    dctFile.Add sFileName, dct
    
End Sub

Private Function ItemIsSection(ByVal i_item As String, _
                               ByRef i_section_name As String) As Boolean

    If Left(i_item, 1) = "[" And Right(i_item, 1) = "]" Then
        ItemIsSection = True
        i_section_name = Mid(i_item, 2, Len(i_item) - 2)
    End If
    
End Function

Private Function ItemIsValue(ByVal i_item As String, _
                             ByRef i_value_name As String, _
                             ByRef i_value As String) As Boolean

    If Left(i_item, 1) <> "[" Then
        ItemIsValue = True
        i_value_name = Split(i_item, "=")(0)
        i_value = Replace(i_item, i_value_name & "=", vbNullString, 1, 1)
    End If
    
End Function

Private Sub Class_Terminate()
    Set FSo = Nothing
    Set dctFile = Nothing
End Sub

Public Property Get FileName() As String:   FileName = sFileName:   End Property

Public Property Let FileName(ByVal s As String)
    Const PROC = "FileName-Let"
    
    If Not IsValidFileName(s) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "The provided file name is not a valid file full name!"
    sFileName = s

End Property

Private Property Get FileString(Optional ByVal f_file_full_name As String, _
                                Optional ByVal f_append As Boolean = False, _
                                Optional ByRef f_split As String = vbCrLf, _
                                Optional ByVal f_exclude_empty As Boolean = False) As String
' ----------------------------------------------------------------------------
' Returns the content of a file (f_file_full_name) as a single string plus the
' records/lines delimiter (f_split) which may be vbCrLf, vbCr, or vbLf.
' ----------------------------------------------------------------------------

    Open f_file_full_name For Input As #1
    FileString = Input$(lOf(1), 1)
    Close #1
    
    Select Case True
        Case InStr(FileString, vbCrLf) <> 0: f_split = vbCrLf
        Case InStr(FileString, vbCr) <> 0:   f_split = vbCr
        Case InStr(FileString, vbLf) <> 0:   f_split = vbLf
    End Select
    
    '~~ Eliminate a trailing eof if any
    If Right(FileString, 1) = VBA.Chr(26) Then
        FileString = Left(FileString, Len(FileString) - 1)
    End If
    
    '~~ Eliminate any trailing split string
    If Right(FileString, Len(f_split)) = f_split Then
        FileString = Left(FileString, Len(FileString) - Len(f_split))
    End If
    If f_exclude_empty Then
        FileString = StringEmptyExcluded(FileString)
    End If
    
End Property

Private Property Let FileString(Optional ByVal f_file_full_name As String, _
                                Optional ByVal f_append As Boolean = False, _
                                Optional ByRef f_split As String = vbCrLf, _
                                Optional ByVal f_exclude_empty As Boolean = False, _
                                         ByVal f_s As String)
' ----------------------------------------------------------------------------
' Writes a string (f_s) with multiple records/lines delimited by a vbCrLf to
' a file (f_file_full_name).
' ----------------------------------------------------------------------------
    
    If f_append _
    Then Open f_file_full_name For Append As #1 _
    Else Open f_file_full_name For Output As #1
    Print #1, f_s
    Close #1
    
End Property

Public Property Get Section() As String:        Section = sSection: End Property

Public Property Let Section(ByVal s As String)
    Const PROC = "Section-Let"
    
    If Not IsValidSectionName(s) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "Not a vaild section name!"
    sSection = s
    
End Property

Public Property Get Sections(Optional ByVal s_file As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' Returns a Dictionary with all named sections (pp_sections) in file (pp_file) -
' provided as full name string or as file object - whereby each section is a
' Dictionary with the section name as the key - in ascending
' order - and a Dictionary of the section's values as item with the value name
' as key and the value as item.
' ----------------------------------------------------------------------------
    Const PROC = "PPsections-Get"
    
    On Error GoTo eh
    Dim vSection    As Variant
    Dim dct         As New Dictionary
    
    If Not IsValidFileName(s_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
        
    '~~ Return all sections
    With dct
        For Each vSection In SectionNames(s_file)
            .Add vSection, Values(v_file:=s_file, v_section:=vSection)
        Next vSection
    End With

xt: Set Sections = KeySort(dct)
    Set dct = Nothing
    Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let Sections(Optional ByVal s_file As String, _
                                      ByVal s_dct As Dictionary)
' ------------------------------------------------------------------------
' Writes the sections provided as Dictionary (pp_dct) to file (pp_file) -
' provided as full name string or as file object. Existing sections/values
' are overwritten new sections/values are added.
' ------------------------------------------------------------------------
    Const PROC = "PPsections-Get"
    
    On Error GoTo eh
    Dim vName       As Variant
    Dim vSection    As Variant
    Dim dctValues   As Dictionary
    Dim sSection    As String
    Dim sName       As String
    Dim fl          As String
    Dim vValue      As Variant
    
    If Not IsValidFileName(s_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    
    For Each vSection In s_dct
        sSection = vSection
        Set dctValues = s_dct(sSection)
        For Each vName In dctValues
            sName = vName
            vValue = dctValues.item(vName)
            Value(v_file:=s_file _
                , v_section:=sSection _
                , v_value_name:=sName _
                 ) = vValue
        Next vName
    Next vSection
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Get Value(Optional ByVal v_value_name As String = vbNullString, _
                          Optional ByVal v_section As String = vbNullString, _
                          Optional ByVal v_file As String = vbNullString) As Variant
' ----------------------------------------------------------------------------
' Reads a value with a specific name from a section from a file organized:
'[section]
'<value-name>=<value>
' ----------------------------------------------------------------------------
    Const PROC  As String = "Value-Get"
    
    On Error GoTo eh
    Dim lResult As Long
    Dim sRetVal As String
    Dim vValue  As Variant
    Dim fl      As String
    
    If Not IsValidFileName(v_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    If Not IsValidSectionName(v_section) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No valid section name provided!"
    If Not IsValidValueName(v_value_name) _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The provided value name is invalid!"
    
    sRetVal = String(32767, 0)
    lResult = GetPrivateProfileString(lpg_ApplicationName:=v_section _
                                    , lpg_KeyName:=v_value_name _
                                    , lpg_Default:="" _
                                    , lpg_ReturnedString:=sRetVal _
                                    , nSize:=Len(sRetVal) _
                                    , lpg_FileName:=v_file _
                                     )
    Value = Left$(sRetVal, lResult)
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Public Property Let Value(Optional ByVal v_value_name As String = vbNullString, _
                          Optional ByVal v_section As String = vbNullString, _
                          Optional ByVal v_file As String = vbNullString, _
                          ByVal v_value As Variant)
' ----------------------------------------------------------------------------
' Writes a value under a given name (v_value_name) into a section (v_section)
' in a file (v_file) organized:
' [section]
' <value-name>=<value>
' ----------------------------------------------------------------------------
    Const PROC = "Value-Let"
        
    On Error GoTo eh
    Dim lChars  As Long
    Dim sValue  As String
    Dim fl      As String
    
    If Not IsValidFileName(v_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    If Not IsValidSectionName(v_section) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No valid section name provided!"
    If Not IsValidValueName(v_value_name) _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The provided value name is invalid!"
    
    Select Case VarType(v_value)
        Case vbBoolean: sValue = VBA.CStr(VBA.CLng(v_value))
        Case Else:      sValue = v_value
    End Select
    
    lChars = WritePrivateProfileString(lpw_ApplicationName:=v_section _
                                     , lpw_KeyName:=v_value_name _
                                     , lpw_String:=v_value _
                                     , lpw_FileName:=v_file)
    If lChars = 0 Then
        MsgBox "System error when writing property" & vbLf & _
               "Section    = '" & v_section & "'" & vbLf & _
               "Value name = '" & v_value_name & "'" & vbLf & _
               "Value      = '" & CStr(v_value) & "'" & vbLf & _
               "Value file = '" & v_file & "'"
    End If
    
    If dctFile.Exists(sFileName) Then
        '~~ Removing the current file's content from the Dictionary
        '~~ enforces re-load with the next value read.
        dctFile.Remove sFileName
    End If
    
xt: Exit Property

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Property

Private Function ArgStrings(Optional ByVal a_v As Variant = Nothing) As Collection
' ----------------------------------------------------------------------------
' Returns the strings provided in argument (a_v) as a Collection of string
' items whereby (a_v) may be not provided, a comma delimited string, a
' Dictionary, or a Collection of string items.
' ----------------------------------------------------------------------------
    Const PROC = "ArgStrings"
    
    On Error GoTo eh
    Dim cll     As New Collection
    Dim dct     As Dictionary
    Dim vString As Variant
    
    Select Case VarType(a_v)
        Case vbObject
            Select Case TypeName(a_v)
                Case "Dictionary"
                    Set dct = a_v
                    With dct
                        For Each vString In dct
                            cll.Add .item(vString)
                        Next vString
                    End With
                    Set ArgStrings = cll
                Case "Collection"
                    Set ArgStrings = a_v
                Case Else: GoTo xt ' likely Nothing
            End Select
        Case vbString
            If a_v <> vbNullString Then
                With cll
                    For Each vString In Split(a_v, ",")
                        .Add VBA.Trim$(vString)
                    Next vString
                End With
                Set ArgStrings = cll
            End If
        Case Is >= vbArray
        Case Else
            Err.Raise AppErr(1), ErrSrc(PROC), "The argument is neither a string, an arry, a Collecton, or a Dictionary!"
    End Select
            
xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function ErrMsg(ByVal err_source As String, _
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
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf mMsg = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
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
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else: ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "clsPrivProfFile." & sProc
End Function

Public Function Exists(Optional ByVal x_folder As String = vbNullString, _
                       Optional ByVal x_file As String = vbNullString, _
                       Optional ByVal x_section As String = vbNullString, _
                       Optional ByVal x_value_name As String = vbNullString, _
                       Optional ByRef x_result_folder As Folder = Nothing, _
                       Optional ByRef x_result_files As Collection = Nothing) As Boolean
' ----------------------------------------------------------------------------
' Universal File System Objects existence check whereby the existence check
' depend on the provided arguments. The function returns TRUE when:
'
' Argument     | TRUE condition (despite the fact not vbNullString)
' -------------| --------------------------------------------------------
' x_folder     | The folder exists, no x_file provided
' x_file       | When no x_folder had been provided the provided x_file
'              | exists. When an x_folder had been provided at least one
'              | or more x_file meet the LIKE criteria and x_section is
'              | not provided
' x_section    | Exactly one file had been passed the existenc check, the
'              | provided section exists and no x_value_name is provided.
' x_value_name | The provided value-name exists - in the existing section
'              | in the one and only existing file.
' ----------------------------------------------------------------------------
    Const PROC  As String = "Exists"
    
    On Error GoTo eh
    Dim sFileName   As String
    Dim fo          As Folder   ' Folder
    Dim sfo         As Folder   ' Sub-Folder
    Dim fl          As File
    Dim queue       As New Collection
    
    Set x_result_files = New Collection

    With FSo
        If Not x_folder = vbNullString Then
            '~~ Folder existence check
            If Not .FolderExists(x_folder) Then GoTo xt
            Set x_result_folder = .GetFolder(x_folder)
            If x_file = vbNullString Then
                '~~ When no x_file is provided, that's it
                Exists = True
                GoTo xt
            End If
        End If
        
        If x_file <> vbNullString And x_folder <> vbNullString Then
            '~~ For the existing folder an x_file argument had been provided
            '~~ This is interpreted as a "Like" existence check is due which
            '~~ by default includes all subfolders
            sFileName = x_file
            Set fo = .GetFolder(x_folder)
            Set queue = New Collection
            queue.Add fo

            Do While queue.Count > 0
                Set fo = queue(queue.Count)
                queue.Remove queue.Count ' dequeue the processed subfolder
                For Each sfo In fo.SubFolders
                    queue.Add sfo ' enqueue (collect) all subfolders
                Next sfo
                For Each fl In fo.Files
                    If VBA.Left$(fl.Name, 1) <> "~" _
                    And fl.Name Like x_file Then
                        '~~ The file in the (sub-)folder meets the search criteria
                        '~~ In case the x_file does not contain any "LIKE"-wise characters
                        '~~ only one file may meet the criteria
                        x_result_files.Add fl
                        Exists = True
                     End If
                Next fl
            Loop
            If x_result_files.Count <> 1 Then
                '~~ None of the files in any (sub-)folder matched with x_file
                '~~ or more than one file matched
                GoTo xt
            End If
        ElseIf x_file <> vbNullString And x_folder = vbNullString Then
            If Not .FileExists(x_file) Then GoTo xt
            x_result_files.Add .GetFile(x_file)
            If x_section = vbNullString Then
                '~~ When no section is provided, that's it
                Exists = True
                GoTo xt
            End If
        End If
        
        '~~ At this point either a provided folder together with a provided file matched exactly one existing file
        '~~ or a specified file's existence had been proved
        If x_section <> vbNullString Then
            If Not SectionExists(s_file:=x_file, s_section:=x_section) Then GoTo xt
            If x_value_name = vbNullString Then
                '~~ When no x_value_name is provided, that's it
                Exists = True
            Else
                Exists = ValueExists(v_file:=x_file, v_section:=x_section, v_value_name:=x_value_name)
            End If
        End If
    End With
        
xt: Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Private Function FileFullName(ByVal f_file As Variant, _
                              ByVal f_service As String, _
                              ByRef f_file_name As String) As String
' ----------------------------------------------------------------------------
' Returns Private Profile File (f_file) as full file name string
' (f_file_name). When the file (f_file) is a vbNullString, a file selection
' dialog is evoked to select one. When finally there's still no file provided
' a vbNullString is returned.
' ----------------------------------------------------------------------------
    Const PROC = "File"
    
    On Error GoTo eh
    Dim fl  As File
    
    If VarType(f_file) = vbObject Then
        If f_file Is Nothing Then
            If Not FilePicked(p_title:="Select a PrivateProfile file not provided for the service '" & f_service & "'" _
                              , p_init_path:=ThisWorkbook.Path _
                              , p_filters:="Config files, *.cfg; Init files, *.ini; Application data files, *.dat; All files, *.*" _
                              , p_file:=fl _
                               ) _
            Then GoTo xt
            f_file_name = fl.Path
            FileFullName = fl.Path
        ElseIf TypeName(f_file) = "File" Then
            f_file_name = f_file.Path
            FileFullName = f_file.Path
            GoTo xt
        End If
    ElseIf VarType(f_file) = vbString Then
        If f_file = vbNullString Then
            If Not FilePicked(p_title:="Select a PrivateProfile file" _
                              , p_init_path:=ThisWorkbook.Path _
                              , p_filters:="Config files, *.cfg; Init files, *.ini; Application data files, *.dat; All files, *.*" _
                              , p_file:=fl _
                               ) _
            Then GoTo xt
            f_file_name = fl.Path
            FileFullName = fl.Path
        Else
            f_file_name = f_file
            FileFullName = f_file
        End If
    End If

xt: Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function IsValidFileName(ByRef s As String)
' ----------------------------------------------------------------------------
' Retunrns TRUE when a provided string (s) is a vaild full file name. I.e. it
' contains a : and a \ as the minimum.
' ----------------------------------------------------------------------------
    If s = vbNullString Then
        s = sFileName
        IsValidFileName = True
    Else
        With FSo
            If InStr(s, ":") = 0 Then Exit Function
            If InStr(s, "\") = 0 Then Exit Function
            If .GetExtensionName(s) = vbNullString Then Exit Function
            If .FileExists(s) Then
                IsValidFileName = True
                sFileName = s
            Else
                On Error Resume Next
                .CreateTextFile s
                If Err.Number = 0 Then
                    IsValidFileName = True
                    sFileName = s
                    .DeleteFile s
                End If
            End If
        End With
    End If
        
End Function

Private Function IsValidSectionName(ByRef i_section As String) As Boolean
    
    If i_section = vbNullString Then
        i_section = sSection
    Else
        With Reg
            .Pattern = "^[a-zA-Z0-9_]+$"
            IsValidSectionName = .Test(i_section)
        End With
    End If
    
End Function

Public Function IsValidValueName(ByVal i_name As String) As Boolean
    With Reg
        .Pattern = "^[a-zA-Z0-9_]{1,32}$"
        IsValidValueName = .Test(i_name)
    End With
End Function

Public Function KeySort(ByRef k_dct As Dictionary) As Dictionary
' ------------------------------------------------------------------------------
' Returns the items in a Dictionary (k_dct) sorted by key.
' ------------------------------------------------------------------------------
    Const PROC  As String = "KeySort"
    
    On Error GoTo eh
    Dim dct     As New Dictionary
    Dim vKey    As Variant
    Dim arr()   As Variant
    Dim Temp    As Variant
    Dim i       As Long
    Dim j       As Long
    
    If k_dct Is Nothing Then GoTo xt
    If k_dct.Count = 0 Then GoTo xt
    
    With k_dct
        ReDim arr(0 To .Count - 1)
        For i = 0 To .Count - 1
            arr(i) = .Keys(i)
        Next i
    End With
    
    '~~ Bubble sort
    For i = LBound(arr) To UBound(arr) - 1
        For j = i + 1 To UBound(arr)
            If arr(i) > arr(j) Then
                Temp = arr(j)
                arr(j) = arr(i)
                arr(i) = Temp
            End If
        Next j
    Next i
        
    '~~ Transfer based on sorted keys
    For i = LBound(arr) To UBound(arr)
        vKey = arr(i)
        dct.Add Key:=vKey, item:=k_dct.item(vKey)
    Next i
    
xt: Set k_dct = dct
    Set KeySort = dct
    Set dct = Nothing
    Exit Function

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub NamesRemove(ByVal n_value_names As Variant, _
              Optional ByVal n_file As String = vbNullString, _
              Optional ByVal n_section As String = vbNullString, _
              Optional ByRef n_removed As Long)
' -----------------------------------------------------------------------------
' PrivateProfile file service. Removes from a PrivateProfile file's (n_file)
' section (n_section) a name (n_value_names) provided either as a
' comma delimited string, or as one string only.
' When no file (n_file) is provided - either as full name or as file object
' a file selection dialog is displayed. When finally there's still no file
' provided the service ends without notice.
' When the name existed and has been removed, the function returns TRUE.
' -----------------------------------------------------------------------------
    Const PROC = "NamesRemove"
    
    On Error GoTo eh
    Dim v   As Variant
    
    If Not IsValidFileName(n_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    If Not IsValidSectionName(n_section) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No valid section name provided!"
    
    For Each v In Split(Replace(n_value_names, " ", vbNullString), ",")
        If ValueExists(n_file, v, n_section) Then
            DeletePrivateProfileKey Section:=n_section _
                                  , Key:=v _
                                  , Setting:=0 _
                                  , Name:=n_file
            n_removed = n_removed + 1
        End If
    Next v
    If n_removed > 0 Then Set dctSections = Null
    
xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub README(Optional ByVal r_bookmark As String = vbNullString)
    
    If r_bookmark = vbNullString Then
        ShellRun GITHUB_REPO_URL
    Else
        r_bookmark = Replace("#" & r_bookmark, "##", "#") ' add # if missing
        ShellRun GITHUB_REPO_URL & r_bookmark
    End If
        
End Sub

Public Sub Reorg(Optional ByVal r_file As String = vbNullString)
' ----------------------------------------------------------------------------
' Private-Profile file service. Re-organizes all sections and all value-names
' therein in ascending order. When no file (r_file) is provided - either as
' full name or as file object a file selection dialog is displayed. When
' finally there's still no file provided the service ends without notice. The
' processed file is returned (r_file_result).
' Restriction: The file must not contain any comment lines other than at the
'              top of the very first section.
' ----------------------------------------------------------------------------
    Const PROC = "Reorg"
    
    On Error GoTo eh
    Dim vSection    As Variant
    Dim sSection    As String
    Dim dctSections As Dictionary
    Dim dctValues   As Dictionary
    Dim vValue      As Variant
    Dim fl          As String
    
    If Not IsValidFileName(r_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    
    If dctSections Is Nothing Then Set dctSections = Sections(r_file)
    For Each vSection In dctSections
        SectionsRemove s_file:=r_file _
                     , s_sections:=vSection _
                     , s_reorg:=False ' prevent Set Null of dctSections
        Set dctValues = dctSections(vSection)
        For Each vValue In dctValues
            Value(v_section:=vSection, v_value_name:=vValue) = dctValues(vValue)
        Next vValue
    Next vSection

xt: Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Function SectionExists(Optional ByVal s_file As String = vbNullString, _
                              Optional ByVal s_section As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' Private Profile file service: Returns TRUE when the section (s_section)
' exists in file (s_file).
' ----------------------------------------------------------------------------
    Const PROC = "SectionExists"
    
    Dim dctF As Dictionary
    
    If Not IsValidFileName(s_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    If Not IsValidSectionName(s_section) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No valid section name provided!"
        
    If Not dctFile.Exists(sFileName) Then
        '~~ Any action which changes the content removes the FileName entry in the dctFile
        '~~ and the first subsequent read has to restore it at first.
        FileAsDict sFileName
    End If
    If dctFile.Exists(sFileName) Then
        Set dctF = dctFile(sFileName)
        SectionExists = dctF.Exists(s_section)
    End If
    
End Function

Public Function SectionNames(Optional ByVal s_file As String) As Dictionary
' ----------------------------------------------------------------------------
' Private Profile File service: Returns a Dictionary with all section names in
' the file (s_file) whereby the file (s_file) may be provided as full name
' string or as a file object. The sections are returned in ascending order.
' ----------------------------------------------------------------------------
    Const PROC = "SectionNames"
    
    On Error GoTo eh
    Dim asSections()    As String
    Dim dct             As New Dictionary
    Dim fl              As String
    Dim iLen            As Long
    Dim strBuffer       As String
    Dim v               As Variant
    Dim s               As String
    
    Set SectionNames = dct
    
    If Not IsValidFileName(s_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
        
    With dct
        For Each v In Split(FileString(sFileName), vbCrLf)
            If ItemIsSection(v, s) Then
                .Add s, s
            End If
        Next v
    End With
    
xt: Set SectionNames = KeySort(dct)
    Set dct = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Sub SectionsCopy(ByVal s_source As String, _
                        ByVal s_target As String, _
               Optional ByVal s_sections As Variant = Nothing, _
               Optional ByVal s_merge As Boolean = False)
' ----------------------------------------------------------------------------
' PrivateProfile file service.
' Copies sections from file (s_source) to file (s_target), when no section
' names (s_sections) are provided all. By default (s_merge) all sections
' are replaced.
' ----------------------------------------------------------------------------
    Const PROC = "PPsectionCopy"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim vName       As Variant
    Dim vSection    As Variant
    
    For Each vSection In ArgStrings(s_sections)
        If Not s_merge Then
            '~~ Section will be replaced
            SectionsRemove s_file:=s_target _
                         , s_sections:=vSection _
                         , s_reorg:=False
        End If
        Set dct = Values(v_file:=s_source _
                       , v_section:=vSection)
        For Each vName In dct
            Value(v_file:=s_target _
                , v_section:=vSection _
                , v_value_name:=vName) = dct(vName)
        Next vName
     Next vSection

xt: Set dct = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub SectionsRemove(Optional ByVal s_sections As String = vbNullString, _
                          Optional ByVal s_file As String = vbNullString, _
                          Optional ByRef s_reorg As Boolean = True)
' ----------------------------------------------------------------------------
' PrivateProfile file service.
' Removes the sections (s_sections) - provided as: comma delimited string,
' Dictionary with name items, or Collection of names, in file (s_file) -
' provided as full name string or as file object.
' ----------------------------------------------------------------------------
    Const PROC = "SectionsRemove"
    
    On Error GoTo eh
    Dim v           As Variant
    Dim sSection    As String
    
    If Not IsValidFileName(s_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    
    For Each v In ArgStrings(s_sections)
        sSection = v
        If IsValidSectionName(sSection) _
        Then DeletePrivateProfileSection Section:=v _
                                       , NoKey:=0 _
                                       , NoSetting:=0 _
                                       , Name:=s_file
    Next v
    If Not s_reorg And dctFile.Exists(s_file) Then
        dctFile.Remove (s_file)
    End If
    
xt: Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Function ShellRun(ByVal s_string As String, _
                 Optional ByVal s_show_how As Long = WIN_NORMAL) As String
' ----------------------------------------------------------------------------
' Opens a folder, email-app, url, or even an Access instance.
'
' Usage Examples: - Open a folder:  ShellRun("C:\TEMP\")
'                 - Call Email app: ShellRun("mailto:user@tutanota.com")
'                 - Open URL:       ShellRun("http://.......")
'                 - Unknown:        ShellRun("C:\TEMP\Test") (will call
'                                   "Open With" dialog)
'                 - Open Access DB: ShellRun("I:\mdbs\xxxxxx.mdb")
' Copyright:      This code was originally written by Dev Ashish. It is not to
'                 be altered or distributed, except as part of an application.
'                 You are free to use it in any application, provided the
'                 copyright notice is left unchanged.
' Courtesy of:    Dev Ashish
' ----------------------------------------------------------------------------

    Dim lRet            As Long
    Dim varTaskID       As Variant
    Dim stRet           As String
    Dim hWndAccessApp   As Long
    
    '~~ First try ShellExecute
    lRet = apiShellExecute(hWndAccessApp, vbNullString, s_string, vbNullString, vbNullString, s_show_how)
    
    Select Case True
        Case lRet = ERROR_OUT_OF_MEM:       stRet = "Execution failed: Out of Memory/Resources!"
        Case lRet = ERROR_FILE_NOT_FOUND:   stRet = "Execution failed: File not found!"
        Case lRet = ERROR_PATH_NOT_FOUND:   stRet = "Execution failed: Path not found!"
        Case lRet = ERROR_BAD_FORMAT:       stRet = "Execution failed: Bad File Format!"
        Case lRet = ERROR_NO_ASSOC          ' Try the OpenWith dialog
            varTaskID = Shell("rundll32.exe shell32.dll,OpenAs_RunDLL " & s_string, WIN_NORMAL)
            lRet = (varTaskID <> 0)
        Case lRet > ERROR_SUCCESS:          lRet = -1
    End Select
    
    ShellRun = lRet & IIf(stRet = vbNullString, vbNullString, ", " & stRet)

End Function

Public Function Value2(ByVal v_value_name As String, _
              Optional ByVal v_section As String = vbNullString, _
              Optional ByVal v_file As String = vbNullString) As Variant
    
    Const PROC = "Value2"
    Dim dctValues   As Dictionary
    Dim dct         As Dictionary
    Dim dctS        As Dictionary
    Dim dctF        As Dictionary
    
    If Not IsValidFileName(v_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    If Not IsValidSectionName(v_section) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No valid section name provided!"
    If Not IsValidValueName(v_value_name) _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The provided value name is invalid!"
        
    If Not dctFile.Exists(sFileName) Then
        '~~ Any action which changes the content removes the FileName entry in the dctFile
        '~~ and the first subsequent read has to restore it at first.
        FileAsDict sFileName
    End If
    If dctFile.Exists(sFileName) Then
        Set dctF = dctFile(sFileName)
        If dctF.Exists(v_section) Then
            Set dctS = dctF(v_section)
            If dctS.Exists(v_value_name) Then
                Value2 = dctS(v_value_name)
            End If
        End If
    End If
    
End Function

Public Function ValueExists(ByVal v_value_name As String, _
                   Optional ByVal v_section As String = vbNullString, _
                   Optional ByVal v_file As String = vbNullString) As Boolean
' ----------------------------------------------------------------------------
' PrivateProfile file service. Returns TRUE when the named value
' (v_value_name) exists in the provided section (v_section) in the file
' (v_file) - provided as full name string or file object. When no file is
' provided a file selection dialog is displayed. When finally there's still no
' file provided the function returns FALSE without notice.
' ----------------------------------------------------------------------------
    Const PROC = "ValueExists"
    
    Dim dctS        As Dictionary
    Dim dctF        As Dictionary
    
    If Not IsValidFileName(v_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    If Not IsValidSectionName(v_section) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No valid section name provided!"
    If Not IsValidValueName(v_value_name) _
    Then Err.Raise AppErr(3), ErrSrc(PROC), "The provided value name is invalid!"
        
    If Not dctFile.Exists(sFileName) Then
        '~~ Any action which changes the content removes the FileName entry in the dctFile
        '~~ and the first subsequent read has to restore it at first.
        FileAsDict sFileName
    End If
    If dctFile.Exists(sFileName) Then
        Set dctF = dctFile(sFileName)
        If dctF.Exists(v_section) Then
            Set dctS = dctF(v_section)
            ValueExists = dctS.Exists(v_value_name)
        End If
    End If

End Function

Public Function ValueNameRename(ByVal v_name_old As String, _
                                ByVal v_name_new As String, _
                                ByVal v_file As Variant, _
                       Optional ByVal v_section As String = vbNullString, _
                       Optional ByRef v_reorg As Boolean = True) As Boolean
' ----------------------------------------------------------------------------
' PrivateProfile file service. Rename an old value name (v_name_old) to a new
' name (v_nmae_new) either in a specific section (v_section) or in all
' sections when no specific section is provided. Optionally not reorgs the
' file when (v_reorg) = False (defaults to True).
' ----------------------------------------------------------------------------
    Const PROC = "ValueNameRename"
    
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim v           As Variant
    Dim vValue      As Variant
    Dim dctSects    As Dictionary
    Dim vSect       As Variant
    
    If v_name_old = vbNullString Or v_name_new = vbNullString Then GoTo xt
    
    Set dct = PPsections(v_file)
    
    For Each v In dct
        If v_section <> vbNullString Then
            If v = v_section Then
                '~~> Rename Name just within one specific section
                If ValueExists(v_file, v_name_old, v_section) Then
                    vValue = Value(v_file, v_section, v_name_old)
                    NamesRemove v_file, v_section, v_name_old
                    If vValue <> vbNullString _
                    Then Value(v_file, v_section, v_name_new) = vValue
                    ValueNameRename = True
                End If
                GoTo xt
            End If
        Else
            Set dctSects = PPsections(v_file)
            For Each vSect In dctSects
            '~~> Rename Name in all sections
                If ValueExists(v_file, v_name_old, v_section) Then
                    vValue = Value(v_file, vSect, v_name_old)
                    NamesRemove v_file, vSect, v_name_old
                    If vValue <> vbNullString Then Value(v_file, vSect, v_name_new) = vValue
                    ValueNameRename = True
                End If
            Next vSect
        End If
    Next v
    
xt: Set dct = Nothing
    If v_reorg Then Reorg v_file
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function ValueNames(Optional ByVal v_file As String = vbNullString, _
                           Optional ByVal v_section As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' PrivateProfile service: Returns a Dictionary of all value names in the
' provided section (v_section), with the value name as key and the value as
' item, in ascending order by value name, in file (v_file) - provided either
' as full file name or as file object.
' ----------------------------------------------------------------------------
    Const PROC = "ValueNames"
    
    On Error GoTo eh
    Dim asNames()   As String
    Dim dct         As New Dictionary
    Dim i           As Long
    Dim lResult     As Long
    Dim sNames      As String
    Dim strBuffer   As String
    Dim v           As Variant
    Dim sSection    As String
    Dim sName       As String
    Dim vNames      As Variant
    Dim fl          As String
    
    If Not IsValidFileName(v_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    If Not IsValidSectionName(v_section) _
    Then Err.Raise AppErr(2), ErrSrc(PROC), "No valid section name provided!"
    
    Set vNames = ArgStrings(v_section)
    If vNames.Count = 0 Then Set vNames = SectionNames(v_file)
    
    For Each v In vNames
        sSection = v
        '~~> Retrieve the names for the provided section
        strBuffer = Space$(32767)
        lResult = GetPrivateProfileString(lpg_ApplicationName:=sSection _
                                        , lpg_KeyName:=vbNullString _
                                        , lpg_Default:=vbNullString _
                                        , lpg_ReturnedString:=strBuffer _
                                        , nSize:=Len(strBuffer) _
                                        , lpg_FileName:=v_file _
                                         )
        sNames = Left$(strBuffer, lResult)
    
        If sNames <> vbNullString Then                                         ' If there were any names
            asNames = Split(sNames, vbNullChar)                      ' have them split into an array
            For i = LBound(asNames) To UBound(asNames)
                sName = asNames(i)
                If Len(sName) <> 0 Then
                    If Not dct.Exists(sName) _
                    Then dct.Add sName, Value(v_file:=fl _
                                                        , v_section:=sSection _
                                                        , v_value_name:=sName)
                End If
            Next i
        End If
    Next v
        
    Set ValueNames = KeySort(dct)

xt: Set dct = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

Public Function Values(Optional ByVal v_section As String = vbNullString, _
                       Optional ByVal v_file As String = vbNullString) As Dictionary
' ----------------------------------------------------------------------------
' PrivateProfile file service: Returns all values in section (v_section) in
' file (v_file) as Dictionary, with the value name as key and the value as
' item, with the value names in asscending order. When the provided section
' does not exist the returned Dictionary is empty.
' ----------------------------------------------------------------------------
    Const PROC = "Values"
    
    On Error GoTo eh
    Dim asNames()   As String
    Dim i           As Long
    Dim lResult     As Long
    Dim sNames      As String
    Dim strBuffer   As String
    Dim sName       As String
    Dim fl          As String
    Dim dct         As New Dictionary
    Dim v           As Variant
    
    Set Values = New Dictionary ' may remain empty though
    
    If Not IsValidFileName(v_file) _
    Then Err.Raise AppErr(1), ErrSrc(PROC), "No valid files name is provided or available!"
    
    If v_section = vbNullString Then GoTo xt
    '~~> Retrieve the names in the provided section
    strBuffer = Space$(32767)
    lResult = GetPrivateProfileString(lpg_ApplicationName:=v_section _
                                    , lpg_KeyName:=vbNullString _
                                    , lpg_Default:=vbNullString _
                                    , lpg_ReturnedString:=strBuffer _
                                    , nSize:=Len(strBuffer) _
                                    , lpg_FileName:=v_file _
                                     )
    sNames = Left$(strBuffer, lResult)

    If sNames <> vbNullString Then                                         ' If there were any names
        With dct
            For Each v In Split(sNames, vbNullChar)                      ' have them split into an array
                If Len(v) <> 0 Then
                    .Add v, Value(v_section:=v_section _
                                , v_value_name:=v)
                End If
            Next v
        End With
    End If
    
xt: Set Values = KeySort(dct)
    Set dct = Nothing
    Exit Function
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Function

