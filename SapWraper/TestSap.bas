Attribute VB_Name = "TestSap"
Option Explicit

Sub Test_Sap_Init()
    
    Dim sapConnString As String
    sapConnString = "Your Sap connection string"

    Dim oSap As SapWraper: Set oSap = New SapWraper
    With oSap

        .Init "PRZ", sapConnString

        If Not .IsConnected Then
            Debug.Print "Not connected"
            Exit Sub

        Else

            Debug.Print "Connected"

        End If

        Dim oSession As GuiSession: Set oSession = .GetSession
        Set oSession = .GetNewSession
        Set oSession = .GetNewSession
        Set oSession = .GetNewSession
        
        '.CloseSession
        .CloseConnection
        

    End With
    
End Sub
