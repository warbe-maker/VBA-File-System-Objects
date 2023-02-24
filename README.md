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
A kind of a universal existence check service with the following syntax:<br>`mFso.Exists([folder], [file], [section], [value-name], [result_folder], [result_files]`)<br>
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
See table above and inline doku.
### Hint for using the PrivateProfile services
When using them in a VB-Project the following scheme is recommendable. Create a dedicated component (Standard Module) in this example named mIni and copy the below code into it. The advantages:
- Each value-name is a dedicated Get/Let property<br>
Example: `mIni.Any = "xxxxx"` to write and `s = mIni.Any` to read
- The PrivateProfile file's name is used only with the internal Private Value Get/Let services
- The sections are hidden by providing unique value-names. 
```vb
Option Explicit
' ---------------------------------------------------------------------------
' Standard Module mIni
' Maintains PrivateProfile file xxxx.ini (IniFileFullName).
' ---------------------------------------------------------------------------
Private Const VNAME_ANY As String = "AnyName"

Public Property Get IniFileFullName() As String
    CompManCfgFileFullName = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "XXXX.ini")
End Property

Public Property Get Any() As String
    Any = Value(VNAME_ANY, "ThisSection")
End Property

Public Property Let Any(ByVal s As String)
    Value(VNAME_ANY, "ThisSection") = s
End Property


Private Property Get Value(Optional ByVal pp_value_name As String, _
                           Optional ByVal pp_section As String) As Variant
' ----------------------------------------------------------------------------
' Returns the value named (pp_value_name) from the section (pp_section) in the
' PrivateProfile file IniFileFullName.
' ----------------------------------------------------------------------------
    Value = mFso.PPvalue(pp_file:=IniFileFullName _
                      , pp_section:=pp_section _
                      , pp_value_name:=pp_value_name _
                       )
End Property

Private Property Let Value(Optional ByVal pp_value_name As String, _
                           Optional ByVal pp_section As String, _
                                    ByVal pp_value As Variant)
' ----------------------------------------------------------------------------
' Writes the value (pp_value) under the name (pp_value_name) into the
' CompManCfgFileFullName.
' ----------------------------------------------------------------------------
    mFso.PPvalue(pp_file:=IniFileFullName _
               , pp_section:=pp_section _
               , pp_value_name:=pp_value_name _
                ) = pp_value
End Property

```

> This _Common Component_ is prepared to function completely autonomously (download, import, use) but at the same time to integrate with my personal 'standard' VB-Project design. See [Conflicts with personal and public _Common Components_][3] for more details.



## Contribution
Any kind of contribution is welcome. Respecting the (more or less obvious) coding principles will be appreciated. The module is available in a [Workbook][4] (public GitHub repository) which includes a complete regression test of all services.

[1]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-File-Services/master/source/mFso.bas
[2]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Directory-Services/master/source/mDct.bas
[3]:https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
[4]:https://github.com/warbe-maker/Common-VBA-File-Sytem-Objects-Services
