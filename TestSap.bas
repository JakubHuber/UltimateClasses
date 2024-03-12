Attribute VB_Name = "TestSap"
Option Explicit

Sub Test_Sap_Init()
    
    Dim sapConnString As String
    sapConnString = "/SAP_CODEPAGE=1100    /FULLMENU SNC_PARTNERNAME=""p:CN=PRZ, OU=SAP-KITS, O=KFR, C=UK"" SNC_QOP=9 /H/srt.kfplc.com/S/3299/M/vprzcs.sap-bau-prod.c.int.gp.kfplc.com/S/3601/G/USERS /UPDOWNLOAD_CP=2"

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
