Attribute VB_Name = "TestErrorGuard"
Option Explicit

Sub Test_ErrorGuard_VisibleProperties()
    
    Dim oGuard As ErrorGuard: Set oGuard = New ErrorGuard
    With oGuard
        .EnableLogErrorsToFile "TestLog", True
        .DeserializeErrors
    End With
    Set oGuard = Nothing
    
End Sub

Sub Test_ErrorGuard_RaiseGuard()
    Dim oGuard As ErrorGuard: Set oGuard = New ErrorGuard
    
    On Error GoTo ErrHandler
    
    Dim oWorkbook As Workbook: Set oWorkbook = Workbooks.Open("blabla")
    
    Debug.Print 1 / 0
    
    oGuard.DisplayErrors
    oGuard.DisplayErrors True
    Debug.Print oGuard.DeserializeErrors
    
    Set oGuard = Nothing
Exit Sub
ErrHandler:
    oGuard.RaiseGuard "Test_ErrorGuard_RaiseGuard"
    Resume Next
End Sub

Sub Test_ErrorGuard_SaveToLog()
    On Error GoTo ErrHandler
    
    Dim oGuard As ErrorGuard: Set oGuard = New ErrorGuard
    With oGuard
    
        .EnableLogErrorsToFile Format(Date, "yyyy-MM-dd"), True
        
        Debug.Print 1 / 0
        
        Dim oTable As ListObject
        oTable.DataBodyRange(1, 1).Value = 1
        
    End With

Exit Sub
ErrHandler:
    oGuard.RaiseGuard "Test_ErrorGuard_SaveToLog"
    Resume Next

End Sub

Sub Test_ErrorGuard_Exceptions()
    Dim oGuard As ErrorGuard: Set oGuard = New ErrorGuard
    
    With oGuard
        .RaiseGuard "Sap", "555", "No authorisation"
        .RaiseGuard "My source", "-23", "Second error", BusinessException
        
        Debug.Print "GuardErrors: " & .GuardErrors.Count
        
    End With
    
End Sub

Sub Test_ErrorGuard_Deserialize()
    Dim oGuard As ErrorGuard: Set oGuard = New ErrorGuard
    
    On Error GoTo ErrHandler
    
    Dim oWorkbook As Workbook: Set oWorkbook = Workbooks.Open("blabla")
    
    Debug.Print oGuard.DeserializeErrors
    
    Debug.Print 1 / 0
    
    Debug.Print oGuard.DeserializeErrors
    
    Err.Raise 55, "Custom Source", "Custome description"
    
    Debug.Print oGuard.DeserializeErrors
    
    Set oGuard = Nothing
    
Exit Sub
ErrHandler:
    oGuard.RaiseGuard "Test_ErrorGuard_RaiseGuard", , , BusinessException
    Resume Next
End Sub
