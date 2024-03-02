VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ErrorDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum EnumErrorCategories
    SystemException
    BusinessException
End Enum

Public ErrorDateTime As Date
Public ErrorCode As Long
Public Message As String
Public Source As String
Public ErrorCategory As EnumErrorCategories

Public Sub DisplayErrorDetail()
    
    Debug.Print "ErrorCode: " & ErrorCode & ", " & _
        "Message: " & Message & ", " & _
        "Source: " & Source & ", " & _
        "Date: " & Format$(ErrorDateTime, "yyyy-MM-dd, hh:mm:ss") & ", " & _
        "Category: " & ErrorCategoryToString
        
End Sub

Public Function ErrorCategoryToString() As String
    Select Case ErrorCategory
        Case EnumErrorCategories.BusinessException: ErrorCategoryToString = "BusinessException"
        Case EnumErrorCategories.SystemException: ErrorCategoryToString = "SystemException"
        Case Else: ErrorCategoryToString = "Category not recognised"
    End Select
End Function


