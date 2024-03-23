VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SapWraper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#End If

Private Const SAPGUI_OBJECT As String = "SAPGUI"
Private Const SAPGUISERVER_OBJECT As String = "SAPGUISERVER"
Private Const SAPLOGON_PATH As String = "C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe"

Private Enum LogonTypesEnum
    NotSpecified = 0
    SAPGUI = 1
    SAPGUISERVER = 2
End Enum

Private Enum EnumExceptions
    SapSIDMissing = -600
    SapGuiAutoNotFound
    ConnectionsNotFound
    CouldNotFindSessionOrAllSessionsAreBusy
    SessionsNumberLimitReached
    UserOrPasswordFieldNotFound
    LoginToSapFailed
End Enum

Private Type TParams
    ErrorGuard As ErrorGuard
    SapApplication As GuiApplication
    SapGuiAuto As Object
    SapSession As GuiSession
    SapConnection As GuiConnection
    Sid As String
    ConnectionString As String
    LogonType As LogonTypesEnum
    IsConnected As Boolean
End Type
Dim this As TParams

Public Sub Init(sapSid As String, Optional sapConnectionString As String, Optional oGuard As ErrorGuard)
    
    If Not oGuard Is Nothing Then Set this.ErrorGuard = oGuard
    
    If Trim$(sapSid) = vbNullString Then
    
        this.ErrorGuard.RaiseGuard "Sap.Init", EnumExceptions.SapSIDMissing, "Sap SID Missing"
        Exit Sub
        
    End If
    
    this.Sid = sapSid
    this.ConnectionString = sapConnectionString
      
    GetGuiApplication
        
End Sub

Public Sub LoginToSap(userName As String, password As String)
    On Error GoTo ErrHandler
    
    If Not this.IsConnected Then
        Debug.Print "Sap is not connected"
        Exit Sub
    End If
    
    Me.GetSession
    
    If this.SapSession Is Nothing Then
        this.ErrorGuard.RaiseGuard TypeName(Me) & ".LoginToSap", CouldNotFindSessionOrAllSessionsAreBusy, "Could Not Find Session Or All Sessions Are Busy"
        Exit Sub
    End If
    
    With this.SapSession.ActiveWindow
        
        Dim userField As GuiTextField: Set userField = .FindById("usr/txtRSYST-BNAME", False)
        Dim passwordField As GuiPasswordField: Set passwordField = .FindById("usr/pwdRSYST-BCODE", False)
        
        If userField Is Nothing Or passwordField Is Nothing Then
            this.ErrorGuard.RaiseGuard TypeName(Me) & ".LoginToSap", UserOrPasswordFieldNotFound, "User Or Password Field Not Found"
            Exit Sub
        End If
        
        userField.Text = userName
        passwordField.Text = password
        
        .SendVKey 0
        
        Dim oStatusBar As Object: Set oStatusBar = .FindById("sbar")
        If oStatusBar.MessageType = "E" Then
            this.ErrorGuard.RaiseGuard TypeName(Me) & ".LoginToSap", LoginToSapFailed, oStatusBar.Text
        End If
        
    End With

Exit Sub
ErrHandler:
    this.ErrorGuard.RaiseGuard TypeName(Me) & ".LoginToSap"

End Sub

Private Sub GetGuiApplication()
    
    GetSapGuiAuto
    
    'Try to create SapGuiAuto if not exists
    If this.LogonType = NotSpecified Then
        
        CreateSapGuiAuto
        
    End If
        
    On Error GoTo ErrHandler
    
    If this.SapGuiAuto Is Nothing Then
        
        this.ErrorGuard.RaiseGuard TypeName(Me) & ".GetGuiApplication", SapGuiAutoNotFound, "SapGuiAuto Not Found"
        Exit Sub
        
    End If
    
    Set this.SapApplication = this.SapGuiAuto.GetScriptingEngine
    
    If this.SapApplication.Connections.Count = 0 Then
        
        If Trim$(this.ConnectionString) <> vbNullString Then
        
            Set this.SapConnection = this.SapApplication.OpenConnectionByConnectionString(this.ConnectionString, True)
            'Wait, Session might take time to open
            Sleep 3000
            
        Else
        
            this.ErrorGuard.RaiseGuard TypeName(Me) & ".GetGuiApplication", ConnectionsNotFound, "Connections Not Found"
            Exit Sub
        
        End If
        
    End If
    
    'GetConnection of not bussy session
    If this.SapConnection Is Nothing Then
        Set this.SapConnection = GetConnectionOfNonBusySession(this.Sid, this.SapApplication)
    End If

    If Not this.SapConnection Is Nothing Then this.IsConnected = True
    
    
Exit Sub
ErrHandler:
    this.ErrorGuard.RaiseGuard TypeName(Me) & ".GetGuiApplication"
End Sub

Private Sub GetSapGuiAuto()
    On Error Resume Next
    
    Set this.SapGuiAuto = GetObject(SAPGUI_OBJECT)
    If Not this.SapGuiAuto Is Nothing Then
    
        this.LogonType = LogonTypesEnum.SAPGUI
        
    Else
    
        Set this.SapGuiAuto = GetObject(SAPGUISERVER_OBJECT)
        If Not this.SapGuiAuto Is Nothing Then
            this.LogonType = LogonTypesEnum.SAPGUI
        End If

    End If
    
End Sub

Private Sub CreateSapGuiAuto()
    
    Shell SAPLOGON_PATH, vbNormalNoFocus
        
    Sleep 3000
    
    Dim timeout As Date: timeout = Now + TimeValue("00:00:10")
    Do
        
        GetSapGuiAuto
        
    Loop Until (Not this.SapGuiAuto Is Nothing) Or (Now > timeout)
    
End Sub

Private Function GetConnectionOfNonBusySession(sapSid As String, oSapApplication As GuiApplication) As GuiConnection
    On Error GoTo ErrHandler
    
    Dim vSession As GuiSession
    Dim vConnection As GuiConnection
    
    For Each vConnection In oSapApplication.Connections

        For Each vSession In vConnection.Sessions

            If Not vSession.Busy Then

                If vSession.Info.SystemName = sapSid Then

                    Set GetConnectionOfNonBusySession = vSession.Parent
                    Exit For

                End If

            End If

        Next

    Next

Exit Function
ErrHandler:
    this.ErrorGuard.RaiseGuard TypeName(Me) & ".GetConnectionOfNonBusySession"
End Function

Public Function GetSession() As GuiSession
    On Error GoTo ErrHandler
    
    If Not this.IsConnected Then
        Debug.Print "Sap is not connected"
        Exit Function
    End If
    
    Dim vSession As GuiSession: Set vSession = GetFirstNotBusySession()
    Set this.SapSession = vSession
    Set GetSession = vSession
    
Exit Function
ErrHandler:
    this.ErrorGuard.RaiseGuard TypeName(Me) & ".GetSession"
End Function

Private Function GetFirstNotBusySession() As GuiSession
    On Error GoTo ErrHandler

    If Not this.IsConnected Then
        Debug.Print "Sap is not connected"
        Exit Function
    End If

    Dim vSession As GuiSession
    Dim vConnection As GuiConnection

    For Each vConnection In this.SapApplication.Connections

        For Each vSession In vConnection.Sessions

            If Not vSession.Busy Then

                If vSession.Info.SystemName = this.Sid Then

                    Set GetFirstNotBusySession = vSession
                    Exit For

                End If

            End If

        Next

    Next

Exit Function
ErrHandler:
    this.ErrorGuard.RaiseGuard "GetFirstNotBusySession"
End Function

Public Function GetNewSession() As GuiSession
    On Error GoTo ErrHandler
    
    If Not this.IsConnected Then
        Debug.Print "Sap is not connected"
        Exit Function
    End If

    'Get base session and create new one from it
    Dim vSession As GuiSession: Set vSession = GetSession()
    Set GetNewSession = CreateNewSession(vSession)
    
Exit Function
ErrHandler:
    this.ErrorGuard.RaiseGuard "CreateNewSession"
End Function

Public Property Get ErrorGuard() As ErrorGuard
    Set ErrorGuard = this.ErrorGuard
End Property

Private Sub Class_Initialize()
    Set this.ErrorGuard = New ErrorGuard
End Sub

Private Sub Class_Terminate()
    Set this.SapSession = Nothing
    Set this.SapApplication = Nothing
    Set this.SapGuiAuto = Nothing
    Set this.ErrorGuard = Nothing
End Sub

Public Property Get SapSession() As GuiSession
    Set SapSession = this.SapSession
End Property

Public Sub CloseSession()
    On Error GoTo ErrHandler
    
    If Not this.SapSession Is Nothing Then
        this.SapSession.Parent.CloseSession this.SapSession.ID
        Set this.SapSession = Nothing
    End If

Exit Sub
ErrHandler:
    this.ErrorGuard.RaiseGuard TypeName(Me) & ".CloseSession"
End Sub

Public Sub CloseConnection()
    On Error GoTo ErrHandler
    
    If Not this.SapConnection Is Nothing Then
        this.SapConnection.CloseConnection
        Set this.SapSession = Nothing
    End If

Exit Sub
ErrHandler:
    this.ErrorGuard.RaiseGuard TypeName(Me) & ".CloseConnection"
End Sub

Private Function CreateNewSession(oSession As GuiSession) As GuiSession
    On Error GoTo ErrHandler
    
    If oSession Is Nothing Then
        this.ErrorGuard.RaiseGuard TypeName(Me) & ".CreateNewSession", CouldNotFindSessionOrAllSessionsAreBusy, "Could Not Find Session Or All Sessions Are Busy"
        Exit Function
    End If
    
    Dim activeConnection As GuiConnection: Set activeConnection = oSession.Parent
    If activeConnection.Sessions.Count = 6 Then
        this.ErrorGuard.RaiseGuard TypeName(Me) & ".CreateNewSession", SessionsNumberLimitReached, "Sessions Number Limit Reached"
        Exit Function
    End If
    
    'Add existing sessions to temporary dictionary
    Dim sessionList As New Dictionary
    Dim allSessions As Long: allSessions = activeConnection.Sessions.Count
    
    Dim vSession As GuiSession: For Each vSession In activeConnection.Sessions
        If Not vSession.Busy Then sessionList.Add vSession.ID, vSession
    Next
    
    oSession.CreateSession
    
    Dim timeout As Date: timeout = Now + TimeValue("00:00:12")
    Do
        Sleep 750
    Loop Until activeConnection.Sessions.Count > allSessions Or Now > timeout
    
    For Each vSession In activeConnection.Sessions
        
        If Not sessionList.Exists(vSession.ID) Then
        
            Set this.SapSession = vSession
            Set CreateNewSession = this.SapSession
            
        End If
        
    Next
    
Exit Function
ErrHandler:
    this.ErrorGuard.RaiseGuard TypeName(Me) & ".CreateNewSession"
End Function

Public Property Get IsConnected() As Boolean
    IsConnected = this.IsConnected
End Property
