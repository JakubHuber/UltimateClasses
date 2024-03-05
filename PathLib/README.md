###  PathLib class. Perform operations on string path. Work with files/folders system. Access/Excel

**References needed:**
- Microsoft Scripting Runtime
- Microsoft ActiveX Data Objects 6.1 Library
- Microsoft VBScript Regular Expressions 5.5

**Files needed to proper work:**
- JsonConverter.bas [VBA-JSON](https://github.com/VBA-tools/VBA-JSON "VBA-JSON")
- ErrorDetails.cls
- ErrorGuard.cls [ErrorGuardFolder](https://github.com/JakubHuber/UltimateClasses/tree/main/ErrorGuard)

**Usage and how it works:**

PathLib is self instanced class inspired by .Net Path class. Is designed to work and perform operations onf files/folders/path/ondrive. It has internal ErrorGuard which is exposed and can be referenced to other objects. It exposes FileSystem object from Scripting Runtime but lack of some funcionalities pushes me to build around helper functions. Most helpful: `FileToBinaryStream, ConcatenatePath, CreateDirectory`.
Examples of properties and possibilities are in TestPathLib.bas

## **Properties:**
PathLib is self instanced no need to create.

|  Type |Name   |Description   |
| :------------ | :------------------------- | :---------------------------------- |
| Property  |` Fso As FileSystemObject ` | Exposes all funcionality of FileSystem object to perform tasks on folder/files/path  |
| Property  |`ErrorGuard As ErrorGuard`   | Exposes ErrorGuard  |
|Function   | `FileToBinaryStream(FilePath As String) As Byte()`  |  Change file to binary stream. Helpful when working with Azure/Sharepoint and other APIs |
|Function   | `CleanPath(sPath As String) As String`  | Clean path from invalid chars. If file name is provided it also removes invalid chars from filepath  |
|Function   |`ConcatenatePath(ParamArray pathElements() As Variant) As String)`   |  Quick way to concatenate path strings. Curently only ParamArray is supported `ConcatenatePath("c:\a", "b", "file name.txt")`.  |
|Sub   | `CreateDirectory(ByVal folderPath As String)`   | Creates phisical path all the way of **folderPath**. It creates whole path if needed. **folderPath** need to start with drive. Each level of path is checked for folder illegal chars. Error is thrown but folder name is not cleaned. |
|Function   | `IsFolderNameValid(folderName As String) As Boolean ` | Checks if **folderName** has invalid chars  |
| Function  | `RemoveInvalidCharsFromFileName(ByVal fileName As String, Optional replaceInvalidChars As Boolean = False, Optional replacement As String = "_") As String `  | Removes invalid chars from file name. If **replaceInvalidChars** is set to true then invalid chars will be replaced with provided or default **replacement** string |
| Function | `RemoveInvalidCharsFromPath(ByVal sPath As String) As String` | Removes invalid chars from path|
| Function | `HasExctension(ByVal sPath As String) As Boolean` | Similar function to .Net Path.HasExctension function. Determines whether the path includes a file name extension|
| Function | `UriPathToLocal(sPath As String) As String` | Changes OneDrive path to local path. NOTE: works fine with ONEDRIVECOMMERCIAL. TODO: Try figure out how [cristianbuse](https://github.com/cristianbuse/VBA-FileTools)https://github.com/cristianbuse/VBA-FileTools solved GetOneDriveLocalPath  |
