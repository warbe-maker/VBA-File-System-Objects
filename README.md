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

| Name                         | Service                                      |
| ---------------------------- | -------------------------------------------- |
| _FilePrivProfNameRemove_     | Removes a named value entry from a given section in a _PrivateProfile File_ file |
|                              | Syntax:<br>`mFso.NameRemove file, section, valuename` |
| _FilePrivProfSections_       | Returns the sections if a _PrivateProfile File_ as Dictionary with the _Section_ as the key and the _Values_ (see below) as the item |
| _FilePrivProfSectionExists_  | Returns True when a given section [.....] exists in a _PrivateProfile File_ file. |
|                | Syntax:<br>`If mFso.SectionExists(file, section) Then ...`|
| _FilePrivProfSectionsCopy_   | Copies named sections from a source _PrivateProfile File_ file to a target _PrivateProfile File_ file optionally replaced or merged in the target file. |
|                              | Syntax:<br>`mFso.SectionsCopy source-file, target-file, [sections], [replace]`|
| _FilePrivProfSectionsRemove_ | Removes named or all sections [.....] from a  _PrivateProfile File_ file |
|                              | Syntax:<br>`mFso.SectionsRemove file, section`
| _FilePrivProfValue_          | Get: Returns the value for a given name from a given section from a _PrivateProfile File_ <br>Syntax: `v = mFso.Value(file, section, name)`|
|                              | Let: write a value with a given name under a given section in a _PrivateProfile File_ <br>Syntax: `mFso.Value(file, section, name) = value`|
| _FilePrivProfValues_         | Returns the values from a given _PrivateProfile File_ under a given _Section_ as Dictionary with the value name as the key and the value as the item.<br>Syntax: `Set dct = mFso.Values(file, section)` |

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
> This _Common Component_ is prepared to function completely autonomously (download, import, use) but at the same time to integrate with my personal 'standard' VB-Project design. See [Conflicts with personal and public _Common Components_][3] for more details.



## Contribution
Any kind of contribution is welcome. Respecting the (more or less obvious) coding principles will be appreciated. The module is available in a [Workbook][4] (public GitHub repository) which includes a complete regression test of all services.

[1]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-File-Services/master/source/mFso.bas
[2]:https://gitcdn.link/cdn/warbe-maker/Common-VBA-Directory-Services/master/source/mDct.bas
[3]:https://warbe-maker.github.io/vba/common/2022/02/15/Personal-and-public-Common-Components.html
[4]:https://github.com/warbe-maker/Common-VBA-File-Sytem-Objects-Services
