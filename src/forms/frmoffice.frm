VERSION 5.00
Begin VB.Form frmrep 
   BackColor       =   &H00E0E0E0&
   Caption         =   "RECEIPT"
   ClientHeight    =   3180
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   4680
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3180
   ScaleWidth      =   4680
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbrec 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4080
      TabIndex        =   20
      Top             =   1440
      Width           =   2295
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   5
      Left            =   4080
      TabIndex        =   19
      Top             =   1920
      Width           =   2295
   End
   Begin VB.CommandButton cmdrec 
      Caption         =   "&Receipt"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   18
      Top             =   6120
      Width           =   1215
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   9600
      TabIndex        =   17
      Tag             =   "AFD"
      Top             =   6000
      Width           =   1215
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4080
      TabIndex        =   15
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   3
      Left            =   4080
      TabIndex        =   14
      Top             =   5520
      Width           =   2295
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   2
      Left            =   4080
      TabIndex        =   13
      Top             =   4560
      Width           =   2655
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Index           =   1
      Left            =   4080
      TabIndex        =   12
      Top             =   3600
      Width           =   2655
   End
   Begin VB.CommandButton cmdAFD 
      Caption         =   "&Edit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   9600
      TabIndex        =   10
      Tag             =   "AFD"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAFD 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   9600
      TabIndex        =   9
      Tag             =   "AFD"
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSC 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   9600
      TabIndex        =   8
      Tag             =   "SC"
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdSC 
      Caption         =   "&Save"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   7
      Tag             =   "SC"
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdAFD 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   9600
      TabIndex        =   6
      Tag             =   "AFD"
      Top             =   1440
      Width           =   1215
   End
   Begin VB.ComboBox cmbbill 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   6360
      TabIndex        =   5
      Top             =   2880
      Width           =   1695
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Index           =   0
      Left            =   4080
      TabIndex        =   11
      Top             =   2880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "RECEIPT NO."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   2640
      TabIndex        =   21
      Top             =   1920
      Width           =   1155
   End
   Begin VB.Line Line5 
      X1              =   3600
      X2              =   7320
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "RECEIPT DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   495
      Left            =   3240
      TabIndex        =   16
      Top             =   240
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00800080&
      FillColor       =   &H000000C0&
      Height          =   5535
      Left            =   9360
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FF00FF&
      X1              =   8400
      X2              =   8400
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FF00FF&
      X1              =   1440
      X2              =   8400
      Y1              =   7320
      Y2              =   7320
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FF00FF&
      X1              =   1440
      X2              =   8400
      Y1              =   1320
      Y2              =   1320
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FF00FF&
      X1              =   1440
      X2              =   1440
      Y1              =   1320
      Y2              =   7320
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " AMOUNT "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   2760
      TabIndex        =   3
      Top             =   5565
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " BILL NO. "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   2880
      TabIndex        =   0
      Top             =   3000
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " OLD BALANCE "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   2280
      TabIndex        =   2
      Top             =   4560
      Width           =   1440
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " RECEIPT  DATE "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   2280
      TabIndex        =   1
      Top             =   3720
      Width           =   1455
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " NEW BALANCE "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   225
      Left            =   2280
      TabIndex        =   4
      Top             =   6720
      Width           =   1470
   End
End
Attribute VB_Name = "frmrep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim optflag As Integer
Private Sub cmdadd_Click()
    optflag = 1
End Sub
Private Sub cmbbill_click()
    
    If optflag = 1 Or optflag = 2 Then
        
        t(0).Text = cmbbill.Text
        
        getBalanceAmount
        
    End If

End Sub
Private Sub getBalanceAmount()

    Dim sql As String
    
    Dim rstemp As New ADODB.Recordset
    
    sql = "SELECT BALANCES FROM BILLMASTER WHERE BILLNUM=" & cmbbill.Text
    
    Set rstemp = cn.Execute(sql)
    
    If recordCheck(rstemp) = True Then
    
        If IsNull(rstemp.Fields(0)) = False Then
            t(2).Text = rstemp.Fields(0)
        Else
            t(2).Text = ""
        End If
    
    End If

    Set rstemp = Nothing

End Sub
Private Sub cmbrec_Click()
    If optflag = 1 Or optflag = 2 Then
        t(5).Text = cmbrec.Text
    End If
    If optflag = 0 Then
        t(5).Text = cmbrec.Text
        findRecord
    End If
End Sub

Private Sub cmdAFD_Click(Index As Integer)

   If Index = 0 Or Index = 1 Then
            
            setSomeFunctionButtons cmdSC, "SC"
            resetSomeFunctionButtons cmdAFD, "AFD"
            
            unlockAllTextBoxes t
            
            t(5).SetFocus
        
        If Index = 0 Then
            resetAllData t
            getReceiptNo
            optflag = 1
        Else
            optflag = 2
        End If
    
 
    Else
    
On Error GoTo databaseErrors
    
        If (t(0).Text = "") Then
            MsgBox "Select Some receipt No", vbInformation
        Else
            cn.Execute "Delete from receipt where billnum='" & t(0).Text & "'"
            
            putDataToComboBox
            
            resetAllData t
        End If
    
    End If
    
databaseErrors:
    databaseError


End Sub
Private Sub getReceiptNo()
    Dim rstemp As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT MAX(RECNO) REC FROM RECEIPT"
    
    
    Set rstemp = cn.Execute(sql)
    
    If recordCheck(rstemp) = True Then
        
        If IsNull(rstemp.Fields(0)) = True Then
            t(5).Text = 1
        Else
            t(5).Text = rstemp.Fields(0) + 1
        End If
    Else
            t(5).Text = 1
    End If

    
    
    
    
    Set rstemp = Nothing

End Sub


Private Sub cmdclose_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdrec_Click()
    If t(0).Text = "" Then
        MsgBox "Choose Some Receipt number"
    Else
        showrepReport t(0).Text
    End If
End Sub

Private Sub cmdSC_Click(Index As Integer)

    Dim rstemp As New ADODB.Recordset
    Dim d(6) As String
    Dim I As Integer
    Dim sql As String
    
    If Index = 0 Then
        
         
        For I = 0 To 5
            If (t(I).Text = "") Then
                d(I) = "NULL"
            Else
                If (I = 0 Or I = 2 Or I = 3 Or I = 4 Or I = 5) Then
                    d(I) = t(I).Text
                Else
                    d(I) = "'" & t(I).Text & "'"
                End If
            End If
        Next

        If optflag = 1 Then
            sql = "INSERT INTO receipt VALUES(" & d(0) & "," & d(1) & "," & d(2) & "," & d(3) & "," & d(4) & "," & d(5) & ")"
        Else
            sql = "UPDATE receipt SET dor=" & d(1) & ",oldbalance=" & d(2) & ",amount=" & d(3) & ",newbalance=" & d(4) & " ,BILLNUM=" & d(0) & " where RECNO=" & d(5)
        End If
    
On Error GoTo databaseErrors
    
        Set rstemp = cn.Execute(sql)
        
        updateDataOfBillMaster
        
        putDataToComboBox
        
    Else
        
        resetAllData t
    
    End If
    
    
databaseErrors:
    
    databaseError
    
    lockAllTextBoxes t
    
    setSomeFunctionButtons cmdAFD, "AFD"
    
    resetSomeFunctionButtons cmdSC, "SC"
    
    
    optflag = 0
    
    Set rstemp = Nothing

End Sub
Private Sub updateDataOfBillMaster()

    Dim sql As String
    
    Dim rstemp As New ADODB.Recordset
    
    
    sql = "UPDATE BILLMASTER SET BALANCES = " & t(4).Text & " WHERE BILLNUM=" & t(0).Text
    

    Set rstemp = cn.Execute(sql)
    
    Set rstemp = Nothing
    
End Sub

Private Sub findRecord()

    Dim sql As String
    Dim rstemp As New ADODB.Recordset
    
    sql = "SELECT * FROM receipt WHERE RECNO='" & t(5) & "'"
    
    Set rstemp = cn.Execute(sql)
    
    If recordCheck(rstemp) = True Then
            
        displayAllRecords t, rstemp
            
    Else
        resetAllData t
        
    End If
    
    Set rstemp = Nothing

End Sub
Private Sub Form_Load() 'www.freestudentprojects.com
    
    putDataToCmbReceipt
    
    putDataToComboBox
    
End Sub
Private Sub putDataToCmbReceipt()

    Dim sql As String
    Dim rstemp As New ADODB.Recordset
    
    sql = "SELECT RECNO FROM RECEIPT"
    
    Set rstemp = cn.Execute(sql)
    
    cmbrec.Clear
    
    If recordCheck(rstemp) = True Then
        loadDataToSingleComboBox rstemp, cmbrec
    End If
    
    Set rstemp = Nothing

End Sub

Private Sub putDataToComboBox()

    Dim rstemp As New ADODB.Recordset
    
    Set rstemp = cn.Execute("select billnum from BILLMASTER")
    
    cmbbill.Clear
    
    
    If recordCheck(rstemp) = True Then
        
        loadDataToSingleComboBox rstemp, cmbbill
        
    End If

End Sub
Private Sub t_LostFocus(Index As Integer)
    If Index = 3 Then
        If optflag = 1 Or optflag = 2 Then
            t(4).Text = t(2).Text - t(3).Text
        End If
    End If
End Sub
