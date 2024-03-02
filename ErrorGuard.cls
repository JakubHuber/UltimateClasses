VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorGuard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type TParameters
    IsRaised As Boolean
    GuardErrors As Collection
    IsLoggingToFileEnabled  As Boolean
    LogFilePath As String
End Type
Private this As TParameters

Public Sub EnableLogErrorsToFile(ByVal logFileName As String, CreateLogInProjectFolder As Boolean)
    
    Const DEFAULT_LOG_FILE_NAME As String = "ErrorGuard.log"
    If Trim$(logFileName) = vbNullString Then logFileName = DEFAULT_LOG_FILE_NAME
        
    If Not PathLib.HasExctension(logFileName) Then logFileName = logFileName & ".log"
    logFileName = PathLib.RemoveInvalidCharsFromFileName(logFileName)
    
    RegisterLogFilePath logFileName, CreateLogInProjectFolder
    
    PathLib.CreateDirectory PathLib.Fso.GetParentFolderName(this.LogFilePath)
    
    this.IsLoggingToFileEnabled = True
    
    Debug.Print "IsLoggingToFileEnabled: " & this.IsLoggingToFileEnabled, "LogPath: " & this.LogFilePath
    
End Sub

Private Sub RegisterLogFilePath(ByVal logFileName As String, CreateLogInProjectFolder As Boolean)
    
    Const FOLDER_VBA_LOGS As String = "VBA logs"
    
    Dim projectPath As String
    
    
    If CreateLogInProjectFolder Then

        'Determine if Excel or Access
        If DetermineApplication = "Microsoft Excel" Then
            
            this.LogFilePath = PathLib.ConcatenatePath(PathLib.UriPathToLocal(ThisWorkbook.Path), FOLDER_VBA_LOGS, logFileName)
            
        Else
            
            this.LogFilePath = PathLib.ConcatenatePath(CurrentProject.Path, FOLDER_VBA_LOGS, logFileName)
            
        End If
    
    Else
        
        this.LogFilePath = PathLib.ConcatenatePath(Environ$("LOCALAPPDATA"), FOLDER_VBA_LOGS, logFileName)
        
    End If
    
End Sub

Private Function DetermineApplication() As String
    Dim oApp As Application: Set oApp = Application
    
    DetermineApplication = oApp.Name
    
    Set oApp = Nothing
    
End Function

Public Sub RaiseGuard(Source As String, Optional ErrorCode As Long, Optional Message As String, Optional ErrorCategory As EnumErrorCategories = 0)
    
    this.IsRaised = True
    
    Dim errorDetail As ErrorDetails: Set errorDetail = New ErrorDetails
    With errorDetail
    
        .Source = Source
        .ErrorDateTime = Now
        
        If ErrorCode = 0 Then
            .ErrorCode = Err.Number
        Else
            .ErrorCode = ErrorCode
        End If
        
        If Message = vbNullString Then
            .Message = Err.Description
        Else
            .Message = Message
        End If
        
        .ErrorCategory = ErrorCategory
    End With
    
    errorDetail.DisplayErrorDetail
    
    this.GuardErrors.Add errorDetail
    
    If this.IsLoggingToFileEnabled Then LogErrorToFile errorDetail
    
End Sub

Private Sub LogErrorToFile(errorDetail As ErrorDetails)
    
    If this.GuardErrors.Count = 0 Then Exit Sub
    
    Dim oStream As TextStream: Set oStream = PathLib.Fso.OpenTextFile(this.LogFilePath, ForAppending, True)
    With oStream
        .WriteLine "User           : " & Environ$("USERNAME")
        .WriteLine "Machine        : " & Environ$("COMPUTERNAME")
        .WriteLine "Current Time   : " & errorDetail.ErrorDateTime
    
        .WriteLine "Source         : " & errorDetail.Source
        .WriteLine "ErrorCode      : " & errorDetail.ErrorCode
        .WriteLine "Message        : " & errorDetail.Message
        .WriteLine "Category       : " & errorDetail.ErrorCategoryToString
        .WriteLine "---------------"
                
    End With
    If Not oStream Is Nothing Then oStream.Close: Set oStream = Nothing
    
End Sub

Public Function DeserializeErrors() As String
    Dim Errors As Collection: Set Errors = New Collection
    Dim Details As Dictionary
    
    Dim errDetail As ErrorDetails: For Each errDetail In this.GuardErrors
        
        Set Details = New Dictionary
        With Details
            .Add "ErrorDateTime", Format$(errDetail.ErrorDateTime, "yyyy-MM-dd hh:mm:ss")
            .Add "Source", errDetail.Source
            .Add "ErrorCode", errDetail.ErrorCode
            .Add "Message", errDetail.Message
            .Add "ErrorCategory", errDetail.ErrorCategoryToString
        End With
        
        Errors.Add Details
        
    Next
    
    DeserializeErrors = JsonConverter.ConvertToJson(Errors)
    
End Function

Private Function ErrorCategoryToString(eErrorCategory As EnumErrorCategories) As String
    Select Case eErrorCategory
        Case EnumErrorCategories.BusinessException: ErrorCategoryToString = "BusinessException"
        Case EnumErrorCategories.SystemException: ErrorCategoryToString = "SystemException"
        Case Else: ErrorCategoryToString = "Category not recognised"
    End Select
End Function

Public Sub DisplayErrors(Optional InMessageBox As Boolean = False)
    
    If this.GuardErrors.Count = 0 Then Exit Sub
    
    Dim errDetail As ErrorDetails
    
    If InMessageBox Then
        
        Dim msg As String
        For Each errDetail In this.GuardErrors
        
            With errDetail
                msg = msg & "[ErrorCode:] " & .ErrorCode & vbTab & _
                    "[Source:] " & .Source & vbTab & _
                    "[Message:] " & .Message & vbNewLine
            End With
        
        Next
        
        MsgBox msg, vbExclamation + vbOKOnly, "Catched errors:"
    
    Else
        
        For Each errDetail In this.GuardErrors
            errDetail.DisplayErrorDetail
        Next
        
    End If

End Sub

Public Sub ClearGuard()
    this.IsRaised = False
    Set this.GuardErrors = New Collection
End Sub

Public Property Get IsRaised() As Boolean
    IsRaised = this.IsRaised
End Property
Public Property Let IsRaised(ByVal vNewValue As Boolean)
    this.IsRaised = vNewValue
End Property

Private Sub Class_Initialize()
    Set this.GuardErrors = New Collection
End Sub

Private Sub Class_Terminate()
    Set this.GuardErrors = Nothing
End Sub

Public Property Get IsLoggingToFileEnabled() As Boolean
    IsLoggingToFileEnabled = this.IsLoggingToFileEnabled
End Property

Public Property Get LogFilePath() As String
    LogFilePath = this.LogFilePath
End Property

Public Property Get GuardErrors() As Collection
    Set GuardErrors = this.GuardErrors
End Property
