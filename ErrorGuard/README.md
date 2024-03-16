###  Error guard classes. Simple idea to catch, manage errors in Access/Excel projects.

**References needed:**
- Microsoft Scripting Runtime
- Microsoft ActiveX Data Objects 6.1 Library
- Microsoft VBScript Regular Expressions 5.5

**Files needed to proper work:**
- JsonConverter.bas [VBA-JSON](https://github.com/VBA-tools/VBA-JSON "VBA-JSON")
- ErrorDetails.cls 
- PathLib.cls [PathLibFolder](https://github.com/JakubHuber/UltimateClasses/tree/main/PathLib)

**Usage and how it works:**

ErrorGuard object is designed to collect all errors encountered during program run. Also it gives output to the Imediate window and log file if needed. It takes properties of global VBA Err object, when error occurs, like message and number if no optional parameters are provided. Minimum to pass to Error Guard is Source. Best in pattern: [Class/Module Name].[Subroutine] (Example `TypeName(Me) & ".MySubroutineName"` )
Put error handler in each subroutine wher you think you can spot for errors or raise your own errors like: oErrorGuard.Raise or Err.Raise 

*** Example implementation in project can be found in Excel file: ErrorGuard process flow example.xlsm. 
Download example and run try step by step function SendMails.
Comments can be search by fraze: NOTE***

Specyfic try-catch procedure template is needed to proper use:
```vb
Public Class

Private Type TConfig
    ErrorGuard As ErrorGuard
End Type
Private this As TConfig

Sub Example
    On Error GoTo ErrHandler

	'Example code . . .

'If needed
Finally:
    'Clean process

Exit Sub

ErrHandler:
    this.ErrorGuard.RaiseGuard TypeName(Me) & ".Example"
    GoTo Finally
End Sub

End Class
```

It is prefered to use it in a class object and reference it if needed to other class objects. 
Also it works good when you want to keep separate error handling for each instance of classes.

Examples of properties possibilities are in TestErrorGuard.bas

## **Properties:**
To start working with ErrorGuard instance it like:
```vb
Dim oGuard as ErrorGuard: Set oGuard = New ErrorGuard
```
|  Type |Name   |Description   |
| :------------ | :------------------------- | :---------------------------------- |
| Property  |` IsRaised As Boolean ` | True when first error show up  |
| Property  |`GuardErrors As Collection`   | Collection of catched errors  |
|Property   | `IsLoggingToFileEnabled  As Boolean`  |  Indicates if errors will be instantly written to file |
|Property   | `LogFilePath As String`  | When logging enabled path will be displayied  |
|Sub   |`EnableLogErrorsToFile(ByVal logFileName As String, CreateLogInProjectFolder As Boolean)`   |Enables loging into text file. Can be pass with or without extension.  If **logFileName** is not provided errors are written to default *ErrorGuard.log*.  **CreateLogInProjectFolder** creates *VBA logs* folder in root project folder. If false then folder is created in user LOCALAPPDATA. Sub displays in Immidiate Window where log will be saved.  |
|Sub   | `Public Sub RaiseGuard(Source As String, Optional ErrorCode As Long, Optional Message As String, Optional ErrorCategory As EnumErrorCategories = 0) `   | Raises guard *IsRaised = True*, adds error to collection, displays error in Immidiate Window and if *IsLoggingToFileEnabled* writes error info to file.  |
|Function   | `DeserializeErrors() As String ` | Convert errors to Json collection  |
| Sub  | `DisplayErrors(Optional InMessageBox As Boolean = False) `  | Best to use at the end of process. Easy way to show errors to user if *InMessageBox* is set to True.     |
| Sub | `ClearGuard()` | Sets *IsRaised* to false and clears GuardErrors collection|
