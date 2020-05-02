VERSION 5.00
Begin VB.Form frmBus 
   BackColor       =   &H00808080&
   Caption         =   "Bus Detail"
   ClientHeight    =   6795
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9480
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
      Left            =   6720
      TabIndex        =   16
      Tag             =   "AFD"
      Top             =   6000
      Width           =   1455
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
      Left            =   9840
      TabIndex        =   8
      Tag             =   "AFD"
      Top             =   1080
      Width           =   1335
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
      Left            =   9840
      TabIndex        =   9
      Tag             =   "SC"
      Top             =   2160
      Width           =   1335
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
      Left            =   9840
      TabIndex        =   10
      Tag             =   "SC"
      Top             =   3240
      Width           =   1335
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
      Left            =   9840
      TabIndex        =   13
      Tag             =   "AFD"
      Top             =   5160
      Width           =   1335
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
      Left            =   9840
      TabIndex        =   11
      Tag             =   "AFD"
      Top             =   4200
      Width           =   1335
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
      Index           =   3
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   5160
      Width           =   1815
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
      Height          =   495
      Index           =   2
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   4200
      Width           =   3015
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
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   3120
      Width           =   3015
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
      Index           =   0
      Left            =   2400
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   1920
      Width           =   3015
   End
   Begin VB.ComboBox cmbBusNo 
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
      ItemData        =   "frmBus.frx":0000
      Left            =   5640
      List            =   "frmBus.frx":0002
      TabIndex        =   14
      Top             =   1920
      Width           =   2295
   End
   Begin VB.ComboBox cmbBusType 
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
      ItemData        =   "frmBus.frx":0004
      Left            =   5640
      List            =   "frmBus.frx":0014
      TabIndex        =   12
      Top             =   3120
      Width           =   2295
   End
   Begin VB.Line Line2 
      X1              =   3240
      X2              =   5880
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "BUS DETAILS"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      TabIndex        =   15
      Top             =   120
      Width           =   3615
   End
   Begin VB.Line Line1 
      X1              =   2760
      X2              =   6120
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Shape Shape1 
      Height          =   5895
      Left            =   9600
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   1815
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   " BUS NO. "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1200
      TabIndex        =   0
      Top             =   1920
      Width           =   840
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   " BUS TYPE "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   1080
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   " NUM OF SEATS "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   600
      TabIndex        =   6
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      Caption         =   " SEAT  TYPE "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   960
      TabIndex        =   4
      Top             =   4200
      Width           =   1125
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000000&
      X1              =   480
      X2              =   480
      Y1              =   720
      Y2              =   6720
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00000000&
      X1              =   8520
      X2              =   8520
      Y1              =   720
      Y2              =   6720
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000000&
      X1              =   480
      X2              =   9240
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Line Line8 
      BorderColor     =   &H00000000&
      X1              =   480
      X2              =   9240
      Y1              =   6720
      Y2              =   6720
   End
End
Attribute VB_Name = "frmBus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim optflag As Integer
Private Sub cmdadd_Click()
    optflag = 1
End Sub
Private Sub cmbBusNo_click()
    t(0).Text = cmbBusNo.Text
    findRecord
End Sub
Private Sub cmbBusType_Click()
    t(1).Text = cmbBusType.Text
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
            MsgBox "Select Some Bus No", vbInformation
        Else
            cn.Execute "Delete from bus where busno='" & t(0).Text & "'"
            
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
    Dim d(4) As String
    Dim I As Integer
    Dim sql As String
    
    If Index = 0 Then
        
        t(1).Text = cmbBusType.Text
    
        For I = 0 To 3
            If (t(I).Text = "") Then
                d(I) = "NULL"
            Else
                If (I = 3) Then
                    d(I) = t(I).Text
                Else
                    d(I) = "'" & t(I).Text & "'"
                End If
            End If
        Next

        If optflag = 1 Then
            sql = "INSERT INTO BUS VALUES(" & d(0) & "," & d(1) & "," & d(2) & "," & d(3) & ")"
        Else
            sql = "UPDATE BUS SET btype=" & d(1) & ",stype=" & d(2) & ",nseat=" & d(3) & " where busno=" & d(0)
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
    
    sql = "SELECT * FROM BUS WHERE BUSNO='" & t(0) & "'"
    
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
    
    Set rstemp = cn.Execute("select busno from bus")
    
    cmbBusNo.Clear
    
    
    If recordCheck(rstemp) = True Then
        
        loadDataToSingleComboBox rstemp, cmbBusNo
        
    End If

End Sub

