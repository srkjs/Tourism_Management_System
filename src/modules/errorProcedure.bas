Attribute VB_Name = "errorProcedure"

Public Sub databaseError()


    Dim e As ADODB.Error
    
    For Each e In cn.Errors
    
        MsgBox e.Description
        
        Select Case getErrorNumber(e.Description)
            Case "00001"
                MsgBox "Duplicate Entry , Please Re-enter", vbInformation
            Case "00984"
                MsgBox "Please Enter Numeric Value", vbInformation
            Case "01401"
                MsgBox "You have entered more no of characters ,Please, check Your Entry "
            Case "02290"
                MsgBox "Please check your entries, Invalid Data Provided"
            Case "01438"
                MsgBox "Please check your entry, too large value provided"
            
       
        End Select
    Next
    
    
    cn.Errors.Clear
    

End Sub
Private Function getErrorNumber(e As String) As String

On Error Resume Next
    
    Dim s As Long
    Dim ee As Long
    
    s = InStr(1, e, "-") + 1
    ee = InStr(1, e, ":")
    
    getErrorNumber = Mid(e, s, ee - s)
    
    


End Function

