VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PathLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Class is preinitialized.

Option Explicit

Private Const BACKSLASH As String = "\"
Private Const SLASH As String = "/"

Private Enum EnumExceptions
    NullException = -800
    InvalidPathChar
    InvalidDriveName
    ObjectsAreNotSupported
End Enum

Private Type TConfig
    Fso As FileSystemObject
    ErrorGuard As ErrorGuard
End Type
Private this As TConfig

Public Function FileToBinaryStream(FilePath As String) As Byte()
    On Error GoTo ErrHandler
    
    If Trim(FilePath) = vbNullString Then Exit Function
    
    Dim oStream As New ADODB.Stream
    With oStream
        .Type = adTypeBinary
        .Open
        .LoadFromFile FilePath
        FileToBinaryStream = .Read
        .Close
    End With
    
Finally:
    Set oStream = Nothing
Exit Function

ErrHandler:
    this.ErrorGuard.RaiseGuard TypeName(Me) & ".FileToBinaryStream"
    On Error GoTo Finally
End Function

Public Function CleanPath(sPath As String) As String
    
    Dim folderPath As String
    Dim fileName As String
    
    If Trim$(sPath) = vbNullString Then Exit Function
    
    If HasExctension(sPath) Then
    
        folderPath = Left(sPath, InStrRev(sPath, BACKSLASH))
        fileName = Right(sPath, Len(sPath) - InStrRev(sPath, BACKSLASH))
        fileName = RemoveInvalidCharsFromFileName(fileName)
        
    Else
        
        folderPath = sPath
        
    End If
    
    folderPath = RemoveInvalidCharsFromPath(folderPath)
    
    CleanPath = this.Fso.BuildPath(folderPath, fileName)
    
End Function

Public Function ConcatenatePath(ParamArray pathElements() As Variant) As String
    
    Dim sPath As String
    Dim i As Long: For i = LBound(pathElements) To UBound(pathElements)
        
        'NOTE: To do - support collection object. Mixed objects?
        If IsObject(pathElements(i)) Then GoTo ThrowObjectsAreNotSupported
        If IsNull(pathElements(i)) Then GoTo ThrowNullException
        If HasPathInvalidChar(CStr(pathElements(i))) Then GoTo ThrowInvalidPathChar
        
        sPath = this.Fso.BuildPath(sPath, CStr(pathElements(i)))
        
    Next
    
    ConcatenatePath = sPath

Exit Function

ThrowNullException:
    this.ErrorGuard.RaiseGuard TypeName(Me) & ".ConcatenatePath", NullException, "Path contains Null"
    Exit Function
    
ThrowInvalidPathChar:
    this.ErrorGuard.RaiseGuard TypeName(Me) & ".ConcatenatePath", InvalidPathChar, "Path contains invalid char"
    Exit Function

ThrowObjectsAreNotSupported:
    this.ErrorGuard.RaiseGuard TypeName(Me) & ".ConcatenatePath", ObjectsAreNotSupported, "Objects are not supported"
    Exit Function

End Function

Public Sub CreateDirectory(ByVal folderPath As String)

    folderPath = Trim$(folderPath)
    If folderPath = vbNullString Then Exit Sub

    If this.Fso.GetDriveName(folderPath) = vbNullString Then
        this.ErrorGuard.RaiseGuard TypeName(Me) & ".CreateDirectory", InvalidDriveName, "Path: " & folderPath & " does not start with drive"
        Exit Sub
    End If

    'If Me.HasExctension(folderPath) Then folderPath = Left(folderPath, InStrRev(folderPath, BACKSLASH) - 1)

    Dim directories() As String: directories = Split(folderPath, BACKSLASH)
    Dim BuildPath As String

    Dim i As Long: For i = LBound(directories) To UBound(directories)

        If i > 0 And Not Me.IsFolderNameValid(directories(i)) Then
            Debug.Print directories(i) & " contains illegal chars!"
            Exit Sub
        End If

        BuildPath = this.Fso.BuildPath(BuildPath, directories(i) & Application.PathSeparator)
        If Not this.Fso.FolderExists(BuildPath) Then this.Fso.CreateFolder BuildPath

    Next

End Sub

Public Function IsFolderNameValid(folderName As String) As Boolean

    Dim invalidCharsPattern As String: invalidCharsPattern = InvalidCharsToRegexPattern

    Dim oRegEx As RegExp: Set oRegEx = New RegExp
    With oRegEx
        .Pattern = "[" & invalidCharsPattern & "]"
        .Global = True
        IsFolderNameValid = Not .Test(folderName)
    End With

End Function

Public Function RemoveInvalidCharsFromFileName(ByVal fileName As String, Optional replaceInvalidChars As Boolean = False, Optional replacement As String = "_") As String

    fileName = Trim$(fileName)

    Dim invalidCharsPattern As String: invalidCharsPattern = InvalidCharsToRegexPattern

    Dim oRegEx As RegExp: Set oRegEx = New RegExp
    With oRegEx
        .Global = True
        .IgnoreCase = True
        .Pattern = "[" & invalidCharsPattern & "]+"
        If replaceInvalidChars Then
            RemoveInvalidCharsFromFileName = .Replace(fileName, replacement)
        Else
            RemoveInvalidCharsFromFileName = .Replace(fileName, vbNullString)
        End If
    End With

End Function

Public Function RemoveInvalidCharsFromPath(ByVal sPath As String) As String
    
    sPath = Trim$(sPath)
    If sPath = vbNullString Then Exit Function
        
    Dim InvalidPathRegex As String: InvalidPathRegex = InvalidPathCharsToRegexPattern
    Dim oRegEx As RegExp: Set oRegEx = New RegExp
    With oRegEx
        .Global = False
        .IgnoreCase = True
        .Pattern = "[" & InvalidPathRegex & "]+"
        RemoveInvalidCharsFromPath = .Replace(sPath, vbNullString)
    End With
    
End Function

Public Function HasExctension(ByVal sPath As String) As Boolean
    
    sPath = Trim$(sPath)
    If sPath = vbNullString Then Exit Function

    HasExctension = Not (this.Fso.GetExtensionName(sPath) = vbNullString)

End Function

Public Function UriPathToLocal(sPath As String) As String

    Const LITERAL_DOCUMENTS As String = "/documents"

    sPath = LCase$(Trim$(sPath))
    If sPath = vbNullString Then Exit Function
    If InStr(sPath, "https://") = 0 Then
        UriPathToLocal = sPath
        Exit Function
    End If

    'Comercial OneDrive has /Documents
    Dim PathOneDriveCommercial As String
    PathOneDriveCommercial = Mid$(sPath, InStr(sPath, LITERAL_DOCUMENTS) + Len(LITERAL_DOCUMENTS), Len(sPath))
    PathOneDriveCommercial = Replace(PathOneDriveCommercial, SLASH, BACKSLASH)

    'TODO: find https path change to drive
    'UriPathToLocal = Replace(Left(Path, InStr(Path, LITERAL_DOCUMENTS) - 1), SLASH, BACKSLASH) & PathOneDriveCommercial


    UriPathToLocal = Environ("ONEDRIVECOMMERCIAL") & PathOneDriveCommercial

End Function

Public Property Get GetInvalidPathChars() As Variant
    
    Dim invalidChars As Variant
    invalidChars = Array("""", "<", ">", "|", _
        Chr(1), Chr(2), Chr(3), Chr(4), Chr(5), Chr(6), Chr(7), Chr(8), Chr(9), Chr(10), _
        Chr(11), Chr(12), Chr(13), Chr(14), Chr(15), Chr(16), Chr(17), Chr(18), Chr(19), Chr(20), _
        Chr(21), Chr(22), Chr(23), Chr(24), Chr(25), Chr(26), Chr(27), Chr(28), Chr(29), Chr(30), Chr(31))
    
    GetInvalidPathChars = invalidChars
    
End Property

Public Function HasPathInvalidChar(sPath As String) As Boolean
    
    If sPath = vbNullString Then Exit Function
        
    Dim InvalidPathRegex As String: InvalidPathRegex = InvalidPathCharsToRegexPattern
    Dim oRegEx As RegExp: Set oRegEx = New RegExp
    With oRegEx
        .Global = False
        .IgnoreCase = True
        .Pattern = "[" & InvalidPathRegex & "]+"
        HasPathInvalidChar = .Test(sPath)
    End With
    
End Function

Public Function InvalidPathCharsToRegexPattern() As String

    Dim invalidChars As Variant: invalidChars = Join(GetInvalidPathChars, vbNullString)
    InvalidPathCharsToRegexPattern = invalidChars

End Function

Public Property Get GetInvalidFileNameChars() As String()

    Const invalidChars As String = "\,<,>,:,"",/,|,?,*"
    GetInvalidFileNameChars = Split(invalidChars, ",")

End Property

Private Function InvalidCharsToRegexPattern() As String

    Dim invalidChars As String: invalidChars = Join(GetInvalidFileNameChars, vbNullString)
    invalidChars = Replace(invalidChars, SLASH, BACKSLASH & SLASH)
    invalidChars = Replace(invalidChars, BACKSLASH, BACKSLASH & BACKSLASH)

    InvalidCharsToRegexPattern = invalidChars

End Function

Private Sub Class_Initialize()
    Set this.Fso = New FileSystemObject
    Set this.ErrorGuard = New ErrorGuard
End Sub

Private Sub Class_Terminate()
    Set this.Fso = Nothing
    Set this.ErrorGuard = Nothing
End Sub

Public Property Get Fso() As FileSystemObject
    Set Fso = this.Fso
End Property

Public Property Get ErrorGuard() As ErrorGuard
    Set ErrorGuard = this.ErrorGuard
End Property

Public Property Set ErrorGuard(ByVal vNewValue As ErrorGuard)
    Set this.ErrorGuard = vNewValue
End Property
