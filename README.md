## VBA File System Object Services
Common services regarding files system objects (files and folders).  The module makes extensive use of the FileSystemObject but goes far beyond the methods provided by this service.
### Universal File/Folder service
| Service                 | Description                                  |
| ----------------------- | -------------------------------------------- |
| _Exists_                | Existence check service See [below](#exists) |

### Common File services

| Service                 | Description                                |
| ----------------------- | ------------------------------------------ |
| _FileArry_              | Get: Returns the content of a text file as an array.|
|                         | Let: Write the content of an array to a file |    
| _FileCompareByWinMerge_ | Function: Displays the differences between two files by means of WinMerge |
| _FileDelete_            | Deletes a file provided either as object or as full name when it exists  |
| _FileDict_              | Returns the content of a test file as Dictionary |
| _FileDiffers_           | Returns a Dictionary of those records which differ between two files providing an _ignore case_ and _ignore empty records_ option |
| _FileExtension_         | Returns a the extension of file's name. The file may be provided either as file object or as full name|
| _FileGet_               | Returns a file object for given file's full name |
| _FileIsValidName_       | Returns TRUE when a provided string conforms with the systems rules for a file name |
| _FilePicked_            | Returns the full name of a file selected in a dialog |
| _FilesSearch_           | Returns a collection of all files found supporting wildcards and sub-folders | 
| _FileTemp_              | Property Get: Provides the full name of an arbitrary named file, by default in the current directory or in a given path with and optional extension which defaults to .tmp | 
| _FileTxt_               | Get: Returns the content of a text file as string, returns the split string/character for the VBA.Split operation which may be used to transfer the string into an array |
|                         | Let: Writes a text string, optionally intermitted by vbCrLf, to a file - optionally appended. |

#### _Exists_ service
A kind of a universal existence check service with the following syntax:<br>`mFso.Exists([ex_folder], [ex_file], [ex_section], [ex_value-name], [ex_result_folder], [ex_result_files]`)<br>
The service has the following named arguments:

| Service              | Description                                |
| -------------------- | ------------------------------------------ |
| _ex\_folder_         | Optional, string expression.<br>The service returns TRUE when the folder exists and no other argument is provided |
| _ex\_file_           | Optional, string expression.<br>When the _ex\_folder_ argument is provided this argument is supposed to be a file name only string which may or may not contain wildcard characters (specification fo a _LIKE_ operator). The function returns any file in any sub-folder which matches the argument string. The function returns TRUE when at least one file matched. When the _ex\_folder_ argument is not provided it is assumed that the argument specifies a full file name and the service returns TRUE when no other arguments are provided |
| _ex\_section_        | Optional, string expression.<br>The service returns TRUE when exactly one existing file matched the above provided arguments and no  _ex\_value\_name_ argument is provided. |
| _ex\_value\_name_    | Optional, string expression.<br>The service returns TRUE when a value with the provide name exists in the provided existing section in the provided existing file  |
| _ex\_result\_folder_ | Optional, Folder expression. Folder object when the _ex\_folder_ argument is an existing folder, else Nothing. |
| _ex\_result\_files_  | Optional, Collection expression.<br>A Collection of file objects with proved  existence |

### _PrivateProfile File_ services
Simplifies the handling of .ini, .cfg, or any other file organized by [section] value-name=value. Consequently all services use the following named arguments:

| Name               | Description                                                                     |
| ------------------ | ------------------------------------------------------------------------------- |
| _pp\_file_         | Variant expression, either a _PrivateProfile File's_ full name or a file object |
| _pp\_section_      | String expression, the name of a _Section_                                      |
| _pp\_value\_name_  | String expression, the name of a _Value_ under a given _Section_                |
| _pp\_value_        | Variant expression, will be written to the file as string                       |

| Name              | Service                                      |
| ------------------| -------------------------------------------- |
| PPsectionExists   |   Returns TRUE when a given section exists in a given PrivateProfile file |
| PPsectionNames    | Returns a Dictionary of all section names [.....] in a PrivateProfile file.
| PPsections        | Get Returns named - or if no names are provided all - sections as Dictionary with the section name as the key and the Values Dictionary as item<br>Let Writes all sections provided as a Dictionary (as above) |
| PPremoveSections  | Removes the sections provided via their name. When no section names are provided (pp_sections) none are removed.
| PPreorg           | Reorganizes all sections and their value-names in a PrivateProfile file by ordering everything in ascending sequence.|
| PPvalue           | Get Reads a named value from a PrivateProfile file<br>Let Writes a named value to a PrivateProfile file |
| PPvalueNameExists | Returns TRUE when a value-name exists in a PrivatProfile file within a given section |
| PPvalueNames      | Returns a Dictionary of all value-names within given sections in a PrivateProfile file with the value-name and the section name as key (<name>[section]) and the value as item, the names in ascending order. Section names may be provided as a comma delimited string. a Dictionary or Collection. When no section names are provided all names of all sections are returned |
| PPvalues          | Returns the value-names and values of a given section in a PrivateProfile file as Dictionary with the value-name as the key (in ascending order) and the value as item.|

### Folder service
| Service                 | Description                                  |
| ----------------------- | -------------------------------------------- |
| _FolderIsValidName_     | Returns TRUE when a provided string conforms with the system's requirements for a correct folder path |
| _Folders_               | Returns all folders in a folder, optionally including all sub-folders, as folder objects in ascending order. When no folder is  provided a folder selection dialog request one. When the provided folder does not exist or no folder is selected the the function returns with an empty  collection. The provided or selected folder is returned as argument. |

## Installation
1. Download and import [mFso.bas][1] to your VB project.
2. In the VBE add a Reference to _Microsoft Scripting Runtime_

## Usage
### PrivateProfile services
The advantages of using the service, specifically when using the below coding scheme:
- Each value-name/value is a dedicated Get/Let property, the values name is used only internally by the service thereby preventing any typos<br>
Write: `mXxxIni.AnyValue = "xxxxx"`<br>
Read:   `Dim sValue As String`<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`sValue = mXxxIni.AnyValue`
Write value under a specific section (in case not hidden by the value-name/value Properties:<br>
Write: `mXxxIni.AnyValue("ThisSection") = "xxxxx"`<br>
Read:   `Dim sValue As String`<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;`sValue = mXxxIni.AnyValue("ThisSection")`
- The PrivateProfile file's name is used only with the internal Private services, e.g. _Value_ Get/Let
- Section names are only exposed when appropriate, e.g. when not included in the name/value properties
- A housekeepingn service ensures that outdated sections are removed (e.g. when a section's name has been changed)

The following recommended scheme provides all the above advantages (when adjusted). Create a dedicated component (Standard Module) for the to-be-maintained ProvateProfile file (mXxxIni in this example) and copy the code below code into it.

```vb
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mIni: Maintains the PrivateProfile file Xxx.ini in
'                       ThisWorkbook's parent folder.
' ---------------------------------------------------------------------------
Private Const VALUE_NAME_ANY    As String = "AnyName"
Private Const SECTION_NAME_ANY  As String = "AnySection"

Public Property Get AnyValue() As String
    AnyValue = Value(VALUE_NAME_ANY, SECTION_NAME_ANY)
End Property

Public Property Let AnyValue(ByVal pp_value As String)
    Value(VALUE_NAME_ANY, SECTION_NAME_ANY) = pp_value
End Property

Private Property Get Value(Optional ByVal pp_value_name As String, _
                           Optional ByVal pp_section As String) As Variant
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' PrivateProfile file XxxIniFullName.
' ----------------------------------------------------------------------------
    Value = mFso.PPvalue(pp_file:=XxxIniFullName _
                      , pp_section:=pp_section _
                      , pp_value_name:=pp_value_name)
End Property

Private Property Let Value(Optional ByVal pp_value_name As String, _
                           Optional ByVal pp_section As String, _
                                    ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' Writes the value (pp_value) under the name (pp_value_name) into the
' PrivateProperty file XxxIniFullName.
' ----------------------------------------------------------------------------
    mFso.PPvalue(pp_file:=XxxIniFullName _
               , pp_section:=pp_section _
               , pp_value_name:=pp_value_name) = pp_value
End Property

Public Property Get XxxIniFullName() As String
    XxxIniFullName = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "Xxx.ini")
End Property

Public Sub Hskpng()
' ------------------------------------------------------------------------------
' Removes obsolete sections, e. g. those with an unknown Name.
' ------------------------------------------------------------------------------
    HskpngRemoveObsoleteSections
  ' HskpngRemoveObsoleteNames (if applicable)
  ' HskpngAddMissingSections (if applicable)
End Sub

Private Sub HskpngRemoveObsoleteSections()
' ------------------------------------------------------------------------------
' Remove sections with an unknown name
' ------------------------------------------------------------------------------
    Dim v As Variant
    
    For Each v In Sections
        If HskpngSectionIsInvalid(v) Then
            RemoveSection v
        End If
    Next v
    
End Sub

Private Function HskpngSectionIsInvalid(ByVal h_section As String) As Boolean
' ------------------------------------------------------------------------------
' Returns TRUE when the section (h_section) is invalid, which is the case when
' it is not valid section's Name.
' ------------------------------------------------------------------------------
    Select Case True
        Case h_section = SECTION_NAME_ANY
        Case Else
            HskpngSectionIsInvalid = True
    End Select
End Function

Private Sub RemoveSection(ByVal r_section_name As String)
    mFso.PPremoveSections pp_file:=XxxIniFullName, pp_sections:=r_section_name
End Sub

Public Sub Reorg()
    mFso.PPreorg XxxIniFullName
End Sub

Private Function Sections() As Dictionary
    Set Sections = mFso.PPsectionNames(XxxIniFullName)
End Function


```
### Other services
#### Exists
Universal File System Objects existence check whereby the existence check depends on the provided arguments. The service has the following named arguments:

| _ex\_folder_ |_ex_file_ |_ex\_section_ |_ex\_value\_name_ |Service returns TRUE when |
|:------------:|:--------:|:------------:|:----------------:|:-------------------------|
| x            |          |              |                  | The folder exists        |
| x            | x        |              |                  | The LIKE file exists in the provided folder |
|              | x        |              |                  | The file exists (needs to be a full path)   |
|              | x        | x            |                  | The section exists in the PrivateProfile file |
|              | x        | x            | x                | The value exists in the provided PrivateProfile file's section |

All existing files are additionally returned in a collection (argument _ex\_result\_files_).
An existing folder is additionally returned as folder object (argument _ex\_result\_folder_).



> This _Common Component_ is prepared to function completely autonomously (download, import, use) but at the same time to integrate with my personal 'standard' VB-Project design. See [Conflicts with personal and public _Common Components_][3] for more details.



## Contribution
Any kind of contribution is welcome. Respecting the (more or less obvious) coding principles will be appreciated. The module is available in a [Workbook][4] (public GitHub repository) which includes a complete regression test of all services.

[1]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-File-Services/master/source/mFso.bas
[2]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Directory-Services/master/source/mDct.bas
[3]:https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
[4]:https://github.com/warbe-maker/Common-VBA-File-Sytem-Objects-Services
