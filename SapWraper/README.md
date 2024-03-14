###  SapWraper - small class to open connection to Sap system and get sessions. Work with Access/Excel.

**References needed:**
- Microsoft Scripting Runtime
- C:\Program Files (x86)\SAP\FrontEnd\SAPgui\sapfewse.ocx

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
More about Sap and GuiConnection object can be found on:
[GuiConnection](https://help.sap.com/docs/sap_gui_for_windows/b47d018c3b9b45e897faf66a6c0885a8/8093f712d0ed4092a463b7edee5792cb.html "GuiConnection")

From this point if property IsConnected is true and you get no errors, you are ready to get GuiSession object using `.GetNewSession` or `.GetSession`

Examples of properties and possibilities are in TestSap.bas

Class defends itself from:
- Sap SID Missing
- SapGuiAuto Not Found 
- Connections Not Found
- Could Not Find Session Or All Sessions Are Busy
- Sessions Number Limit Reached

## **Properties:**
Initialize class like below:
```vb
    Dim oSap As SapWraper: Set oSap = New SapWraper
    oSap.Init "Your Sap SID", "Optional Connection String"
```


|  Type |Name   |Description   |
| :------------ | :------------------------- | :---------------------------------- |
| Property  |`IsConnected() As Boolean` | Shows true if class was able to get SapGuiAuto, SapApplication, GuiConnection |
| Property  |`ErrorGuard As ErrorGuard`   | Exposes ErrorGuard  |
| Sub   | `Init(sapSid As String, Optional sapConnectionString As String, Optional oGuard As ErrorGuard)`  |  If you pass `sapSid` like PRZ/QRZ/FRZ ect. you will get connection of non busy session and you are ready to get Sap session. But you need to have SAPLogon with transaction opened. If you pass optional `sapConnectionString` class will try to open SAPLogon if it is not opened and then connect using sapConnectionString. From that point you are ready to get Sap session.  |
|Function   | `GetSession() As GuiSession`  | Get first session of Sap connection  |
|Function   |`GetNewSession() As GuiSession`   | Creates new session from first non busy session  |
|Sub   |`CloseSession()`   | Closes session which was retrived during `GetSession` or `GetNewSession`  |
|Sub   |`CloseConnection()`   | Closes connection  |


