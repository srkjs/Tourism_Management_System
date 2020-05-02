Attribute VB_Name = "mainProcedure"
Public fMainForm As frmMain

Public cn As ADODB.Connection

Sub Main()
    Dim fLogin As New frmLogin
    
    buildConnection
    
    fLogin.Show vbModal
    If Not fLogin.OK Then
        'Login Failed so exit app
        End
    End If
    Unload fLogin


    frmSplash.Show
    frmSplash.Refresh
    Set fMainForm = New frmMain
    Load fMainForm
    Unload frmSplash


    fMainForm.Show
    
End Sub
Private Sub buildConnection()

    Set cn = New ADODB.Connection
    
    With cn
        .ConnectionString = "Provider=MSDAORA.1;Password=tiger;User ID=scott;Persist Security Info=True"
        .Open
    End With

End Sub
