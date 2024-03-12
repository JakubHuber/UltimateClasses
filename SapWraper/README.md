###  SapWraper - small class to open connection to Sap system and get sessions. Work with Access/Excel.

**References needed:**
- Microsoft Scripting Runtime

**Files needed to proper work:**
- JsonConverter.bas [VBA-JSON](https://github.com/VBA-tools/VBA-JSON "VBA-JSON")
- ErrorDetails.cls
- ErrorGuard.cls [ErrorGuardFolder](https://github.com/JakubHuber/UltimateClasses/tree/main/ErrorGuard)

**Usage and how it works:**

Idea of this small class is to enclose everything what is needed to focus only to get and work with active/new Sap session. It instances all object from SapGuiAuto, GuiApplication, GuiConnection  and GuiSession.  

It can work in two ways.
1. Make user to open sap logon and session. Instance class with Sap SID (like PRZ/QRZ/FRZ ect.) to connect to non busy session or create a new one. Personaly prefer second way as it does not take away working Sap session from user.
```vb
    Dim oSap As SapWraper: Set oSap = New SapWraper
    oSap.Init "PRZ"
```

2.  Instance class with Sap SID and Sap GuiConnection ConnectionString so user will not require to open even Sap logon pad. Sap connection string can be found in GuiConnection.ConnectionString property. It can work best if company uses Auto Sap logon. If not you need to create subroutine to pass user password and login if you want this way.
```vb
Dim sapConnString As String
    sapConnString = "Connection string can be found in GuiConnection.ConnectionString"

    Dim oSap As SapWraper: Set oSap = New SapWraper
    With oSap

        .Init "PRZ", sapConnString
    End with
```

From this point if property IsConnected is true and you get no errors, you are ready to get GuiSession object using .GetNewSession or .GetSession

#TODO:  finish readme

More about Sap and Scripting objects can be found on:
[GuiConnection](https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/8093f712d0ed4092a463b7edee5792cb.html "GuiConnection")

Examples of properties and possibilities are in TestSap.bas


