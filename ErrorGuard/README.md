###  Error guard classes to manage errors in Access/Excel projects.

**References needed:**
- Microsoft Scripting Runtime
- Microsoft ActiveX Data Objects 6.1 Library
- Microsoft VBScript Regular Expressions 5.5

**Files needed to proper work:**
- JsonConverter.bas [VBA-JSON](https://github.com/VBA-tools/VBA-JSON "VBA-JSON")
- PathLib.cls

**Usage:**
ErrorGuard object is designed to collect all errors encountered during program run. Also it gives output to the Imediate window and log file if needed. Specyfic procedure template is needed to proper use:

```vb
Private Type TConfig
    ErrorGuard As ErrorGuard
End Type
Private this As TConfig

Sub Example
    On Error GoTo ErrHandler

	Example code . . .

Exit Sub

ErrHandler:
    this.ErrorGuard.RaiseGuard "Example"
End Sub
```

It is prefered to use it in a class object and reference it to other class objects. 


Test examples of possibilities are in TestErrorGuard.bas
