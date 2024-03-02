Attribute VB_Name = "TestPathLib"
Option Explicit

Sub Test_Path_Uri()
    
    With PathLib
        
        Dim oneDrivePath As String
        oneDrivePath = "https://kfplc-my.sharepoint.com/personal/elzbieta_zielinska_kingfisher_com/Documents/C&B/2023-10_PROMOTIONS AND SALARY UPLIFT"
        Debug.Print .UriPathToLocal(oneDrivePath) = "C:\Users\palase01\OneDrive - Kingfisher PLC\Zielinska, Elzbieta - Zielinska, Elzbieta's files\C&B\2023-10_PROMOTIONS AND SALARY UPLIFT"
        
        'C:\Users\palase01\OneDrive - Kingfisher PLC\Zielinska, Elzbieta - Zielinska, Elzbieta's files\C&B\2023-10_PROMOTIONS AND SALARY UPLIFT\bin
        Debug.Print .UriPathToLocal(ThisWorkbook.Path)
        
        
    End With
    
End Sub

Sub Test_Path_Concatenate()
    
    With PathLib
        Debug.Print .ConcatenatePath("c:\a", "b", "file name.txt")
        Debug.Print .ConcatenatePath("a/", "b?", "file name.txt")
    End With
    
End Sub

Sub Test_Path_CreateDirectory()
    Dim folderPath As String
    
    With PathLib
        
         folderPath = .UriPathToLocal(ThisWorkbook.Path) & "\f1\f2"
        .CreateDirectory folderPath
        
        folderPath = .UriPathToLocal(ThisWorkbook.Path) & "\f3\f4\2023.23.01"
        .CreateDirectory folderPath
        
        'wrong path
        folderPath = "\f3\f4"
        .CreateDirectory folderPath
        
    End With
    
End Sub

Sub Test_Path_HasExctension()
    
    Dim localPath As String
    
    With PathLib
        
        localPath = .UriPathToLocal(ThisWorkbook.Path) & "\f3\f4\file.txt"
        Debug.Print .HasExctension(localPath)
        
        localPath = .UriPathToLocal(ThisWorkbook.Path) & "\f1\f2"
        Debug.Print .HasExctension(localPath)
    
        localPath = .UriPathToLocal(ThisWorkbook.Path) & "\f3\f4\file."
        Debug.Print .HasExctension(localPath)
    
    End With

End Sub

Sub Test_Path_IsFolderNameValid()
    
    Dim localPath As String
    
    With PathLib
        
        Debug.Print "IsFolderNameValid: " & .IsFolderNameValid("f3")
        Debug.Print "IsFolderNameValid: " & .IsFolderNameValid("file.txt")
        Debug.Print "IsFolderNameValid: " & .IsFolderNameValid("f4\file.txt")
        Debug.Print "IsFolderNameValid: " & .IsFolderNameValid("\f3f4")
        
    End With
    
End Sub

Sub Test_Path_RemoveInvalidCharsFromFileName()
    
    Dim localPath As String
    
    With PathLib
        Debug.Print .RemoveInvalidCharsFromFileName("ab\c%234|?.txt")
        Debug.Print .RemoveInvalidCharsFromFileName("ab<c%234|?.txt", True)
        Debug.Print .RemoveInvalidCharsFromFileName("ab>c%234|?.txt", True, "$")
    End With

End Sub
