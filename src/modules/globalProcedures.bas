Attribute VB_Name = "globalProcedures"
Public Function recordCheck(r As ADODB.Recordset) As Boolean

    If r.EOF = True And r.BOF = True Then
        recordCheck = False
    Else
        recordCheck = True
    End If
End Function
Public Sub showBalanceReport()

    rptEnvironment.cmdBalanceAmount_Grouping
    
    If rptEnvironment.rscmdBalanceAmount_Grouping.RecordCount > 0 Then
        rptBalanceReport.Show
    Else
        MsgBox "No Balances"
        rptEnvironment.rscmdBalanceAmount_Grouping.Close
    End If


End Sub

Public Sub unlockAllTextBoxes(t As Object)

    Dim d As TextBox
    
    For Each d In t
        d.Locked = False
    Next

End Sub

Public Sub lockAllTextBoxes(t As Object)

    Dim d As TextBox
    
    For Each d In t
        d.Locked = True
    Next

End Sub
Public Sub resetAllData(t As Object)

    Dim d As TextBox
    
    For Each d In t
        d.Text = ""
    Next
End Sub
Public Sub resetSomeFunctionButtons(t As Object, id As String)
    Dim d As CommandButton
    
    For Each d In t
        If d.Tag = id Then
            d.Enabled = False
        End If
    Next
End Sub
Public Sub setSomeFunctionButtons(t As Object, id As String)
    Dim d As CommandButton
    
    For Each d In t
        If d.Tag = id Then
            d.Enabled = True
        End If
    Next
End Sub
Public Sub displayAllRecords(t As Object, r As Recordset)

    Dim d As TextBox
    Dim i As Integer
    
    
    For Each d In t
        If IsNull(r.Fields(i)) = True Then
            d.Text = ""
        Else
            d.Text = r.Fields(i)
        End If
        i = i + 1
    Next
    
End Sub
Public Sub loadDataToSingleComboBox(r As ADODB.Recordset, cmb As ComboBox)

    Dim i As Integer
    
    cmb.Clear
    
    Do Until r.EOF = True
        cmb.AddItem r.Fields(0)
        r.MoveNext
    Loop

End Sub
Public Sub loadDataToDoubleComboBox(r As ADODB.Recordset, cmb As ComboBox, c As ComboBox)

    Dim i As Integer
    
    cmb.Clear
    c.Clear
    
    Do Until r.EOF = True
        cmb.AddItem r.Fields(0)
        c.AddItem r.Fields(1)
        r.MoveNext
    Loop

End Sub
Public Sub destroyAllForms()



End Sub
Public Sub showBillReport(s As String)
    
    rptEnvironment.cmdBillmaster s
    
    If rptEnvironment.rscmdBillMaster.RecordCount > 0 Then
        rptBill.Show
    Else
        MsgBox "No such BIll"
        rptEnvironment.rscmdBillMaster.Close
    End If

End Sub
Public Sub showTOURReport(s As String)
    
    rptEnvironment.cmdtourpackage
    
    If rptEnvironment.rscmdtourpackage.RecordCount > 0 Then
        rptTour.Show
    Else
        MsgBox "No information found for the request"
        rptEnvironment.rscmdtourpackage.Close
    End If

End Sub
Public Sub showpackReport(s As String)
    
    rptEnvironment.cmdpackages s
    
    If rptEnvironment.rscmdpackages.RecordCount > 0 Then
        rptPackageDetail.Show
    Else
        MsgBox "No such Packages"
        rptEnvironment.rscmdpackages.Close
    End If

End Sub

Public Sub showrepReport(s As String)
    
    rptEnvironment.cmdrep s
    
    If rptEnvironment.rscmdrep.RecordCount > 0 Then
        rptreceiptdetail.Show
    Else
        MsgBox "No such Receipt"
        rptEnvironment.rscmdrep.Close
    End If

End Sub
