VERSION 5.00
Begin VB.Form frmBranch 
   BackColor       =   &H00808080&
   Caption         =   "Branch"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "Times New Roman"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H0000C0C0&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   375
      Index           =   3
      Left            =   8520
      TabIndex        =   13
      Tag             =   "AFD"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.ComboBox cmbbranchno 
      Height          =   405
      ItemData        =   "frmBranch.frx":0000
      Left            =   4080
      List            =   "frmBranch.frx":0002
      TabIndex        =   10
      Top             =   1800
      Width           =   3135
   End
   Begin VB.TextBox t 
      Height          =   375
      Index           =   0
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox t 
      Height          =   1185
      Index           =   1
      Left            =   4080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3000
      Width           =   3135
   End
   Begin VB.TextBox t 
      Height          =   465
      Index           =   2
      Left            =   4080
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4560
      Width           =   3135
   End
   Begin VB.CommandButton cmdAFD 
      Caption         =   "&Edit"
      Height          =   375
      Index           =   1
      Left            =   5400
      TabIndex        =   9
      Tag             =   "AFD"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAFD 
      Caption         =   "&Delete"
      Height          =   375
      Index           =   2
      Left            =   6960
      TabIndex        =   11
      Tag             =   "AFD"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSC 
      Caption         =   "&Cancel"
      Enabled         =   0   'False
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   8
      Tag             =   "SC"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdSC 
      Caption         =   "&Save"
      Enabled         =   0   'False
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   7
      Tag             =   "SC"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.CommandButton cmdAFD 
      Caption         =   "&Add"
      Height          =   375
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Tag             =   "AFD"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   3600
      X2              =   6960
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "BRANCH DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   12
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " TELEPHONE "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   2640
      TabIndex        =   4
      Top             =   4560
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " ADDRESS "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   2760
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " BRANCH NAME "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   225
      Left            =   2280
      TabIndex        =   0
      Top             =   2280
      Width           =   1485
   End
End
Attribute VB_Name = "frmBranch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim optflag As Integer
Private Sub cmdadd_Click()
    optflag = 1
End Sub
Private Sub cmbbranchNo_click()
 t(0).Text = cmbbranchno.Text
 findRecord
End Sub
Private Sub cmdAFD_Click(Index As Integer)

   If Index = 0 Or Index = 1 Then
            
            setSomeFunctionButtons cmdSC, "SC"
            resetSomeFunctionButtons cmdAFD, "AFD"
            
            unlockAllTextBoxes t
            
            t(0).SetFocus
        
        If Index = 0 Then
            resetAllData t
            optflag = 1
        Else
            optflag = 2
        End If
    
 
    Else
    
On Error GoTo databaseErrors
    
        If (t(0).Text = "") Then
            MsgBox "Select Some branch No", vbInformation
        Else
            cn.Execute "Delete from branch where pname='" & t(0).Text & "'"
            
            putDataToComboBox
            
            resetAllData t
        End If
    
    End If
    
databaseErrors:
    databaseError


End Sub

Private Sub cmdclose_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdSC_Click(Index As Integer)

    Dim rstemp As New ADODB.Recordset
    Dim d(3) As String
    Dim I As Integer
    Dim sql As String
    
    If Index = 0 Then
        
       
    
        For I = 0 To 2
            If (t(I).Text = "") Then
                d(I) = "NULL"
            Else
                If (I = 2) Then
                    d(I) = t(I).Text
                Else
                    d(I) = "'" & t(I).Text & "'"
                End If
            End If
        Next

        If optflag = 1 Then
            sql = "INSERT INTO branch VALUES(" & d(0) & "," & d(1) & "," & d(2) & ")"
        Else
            sql = "UPDATE branch SET addres1=" & d(1) & ",tele=" & d(2) & " where pname=" & d(0)
        End If
    
On Error GoTo databaseErrors
    
        Set rstemp = cn.Execute(sql)
        
        putDataToComboBox
        
    Else
        
        resetAllData t
    
    End If
    
    
databaseErrors:
    
    databaseError
    
    lockAllTextBoxes t
    
    setSomeFunctionButtons cmdAFD, "AFD"
    
    resetSomeFunctionButtons cmdSC, "SC"
    
    
    Set rstemp = Nothing
    
    optflag = 0

End Sub
Private Sub findRecord()

    Dim sql As String
    Dim rstemp As New ADODB.Recordset
    
    sql = "SELECT * FROM branch WHERE pname='" & t(0) & "'"
    
    Set rstemp = cn.Execute(sql)
    
    If recordCheck(rstemp) = True Then
            
        displayAllRecords t, rstemp
            
    Else
        
    End If
End Sub
Private Sub Form_Load() 'www.freestudentprojects.com
    
      putDataToComboBox
    
End Sub
Private Sub putDataToComboBox()

    Dim rstemp As New ADODB.Recordset
    
    Set rstemp = cn.Execute("select pname from branch")
    
    cmbbranchno.Clear
    
     If recordCheck(rstemp) = True Then
        
        loadDataToSingleComboBox rstemp, cmbbranchno
        
    End If

End Sub

