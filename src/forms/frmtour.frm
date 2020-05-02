VERSION 5.00
Begin VB.Form frmtour 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   " ORGANIZATION OF TOUR "
   ClientHeight    =   6765
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   7785
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   7785
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbpkiiid 
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
      Left            =   5160
      TabIndex        =   28
      Text            =   "Combo1"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.ComboBox cmbpktype 
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
      ItemData        =   "frmtour.frx":0000
      Left            =   2880
      List            =   "frmtour.frx":000D
      TabIndex        =   27
      Text            =   "Combo"
      Top             =   1320
      Width           =   2175
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   25
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdtour 
      Caption         =   "&Tour"
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
      Left            =   5160
      TabIndex        =   23
      Tag             =   "AFD"
      Top             =   6600
      Width           =   1815
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
      Left            =   8160
      TabIndex        =   22
      Tag             =   "AFD"
      Top             =   6600
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
      Left            =   8160
      TabIndex        =   21
      Tag             =   "AFD"
      Top             =   1800
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
      Left            =   8160
      TabIndex        =   20
      Tag             =   "SC"
      Top             =   2760
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
      Left            =   8160
      TabIndex        =   19
      Tag             =   "SC"
      Top             =   3720
      Width           =   1215
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
      Left            =   8160
      TabIndex        =   18
      Tag             =   "AFD"
      Top             =   4680
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
      Index           =   3
      Left            =   8160
      TabIndex        =   17
      Tag             =   "AFD"
      Top             =   5760
      Width           =   1215
   End
   Begin VB.ComboBox cmbbusno 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5160
      TabIndex        =   16
      Text            =   "COMBO"
      Top             =   5160
      Width           =   1815
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   6
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6600
      Width           =   2055
   End
   Begin VB.ComboBox cmbtid 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5160
      TabIndex        =   15
      Text            =   "COMBO"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   2040
      Width           =   2055
   End
   Begin VB.ComboBox cmbpid 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5160
      TabIndex        =   14
      Text            =   "COMBO"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   5
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   5880
      Width           =   2055
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   4
      Left            =   2880
      Locked          =   -1  'True
      TabIndex        =   8
      Top             =   5160
      Width           =   2055
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   3
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
      Top             =   4440
      Width           =   2055
   End
   Begin VB.TextBox t 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   2
      Left            =   2880
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox t2 
      Height          =   735
      Index           =   0
      Left            =   8040
      TabIndex        =   13
      Top             =   -4920
      Width           =   2535
   End
   Begin VB.Line Line1 
      X1              =   2520
      X2              =   6240
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tour Oraganization "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF00FF&
      Height          =   465
      Left            =   2565
      TabIndex        =   29
      Top             =   120
      Width           =   3645
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " PACKAGE TYPE"
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
      TabIndex        =   26
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " Rs "
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
      Left            =   2520
      TabIndex        =   24
      Top             =   6720
      Width           =   315
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "  FARE  "
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
      Left            =   1680
      TabIndex        =   11
      Top             =   6720
      Width           =   660
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " TOUR ID "
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
      Left            =   1560
      TabIndex        =   0
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " TOTAL SEATS "
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
      TabIndex        =   9
      Top             =   5880
      Width           =   1380
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
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
      Left            =   1560
      TabIndex        =   7
      Top             =   5160
      Width           =   840
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " ARRIVAL "
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
      Left            =   1440
      TabIndex        =   5
      Top             =   4440
      Width           =   945
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " DEPA RTURE "
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
      TabIndex        =   3
      Top             =   3600
      Width           =   1245
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   " PACKAGE  ID "
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
      TabIndex        =   2
      Top             =   2880
      Width           =   1275
   End
End
Attribute VB_Name = "frmtour"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim optflag As Integer
Private Sub cmdadd_Click()
    optflag = 1
End Sub
Private Sub cmbBusNo_click()
    t(4).Text = cmbBusNo.Text
End Sub
Private Sub cmbPkiiid_Click()
    
    If optflag = 1 Or optflag = 2 Then
        t(1).Text = cmbPkiiid.Text
    End If
   
End Sub

Private Sub cmbpktype_Click()
    Dim sql As String
    Dim rstemp As New ADODB.Recordset
    
    sql = "SELECT pkid FROM packages where pkname='" & cmbpktype.Text & "'"
  
    Set rstemp = cn.Execute(sql)
    
    If recordCheck(rstemp) = True Then
 
        
        loadDataToSingleComboBox rstemp, cmbPkiiid
   
    Else
        
    End If
Set rstemp = Nothing
End Sub
Private Sub cmbtid_click()
    t(0).Text = cmbtid.Text
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
            
            GETTOURID
            
        Else
            optflag = 2
        End If
    
 
    Else
    
On Error GoTo databaseErrors
    
        If (t(0).Text = "") Then
            MsgBox "Select Some tour No", vbInformation
        Else
            cn.Execute "Delete from tour where tid='" & t(0).Text & "'"
            
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
    Dim d(7) As String
    Dim I As Integer
    Dim sql As String
    
    If Index = 0 Then
        
       t(1).Text = cmbPkiiid.Text
  
        For I = 0 To 6
            If (t(I).Text = "") Then
                d(I) = "NULL"
            Else
                If (I = 1 Or I = 5 Or I = 6) Then
                    d(I) = t(I).Text
                Else
                    d(I) = "'" & t(I).Text & "'"
                End If
            End If
        Next

        If optflag = 1 Then
            sql = "INSERT INTO tour VALUES(" & d(0) & "," & d(1) & "," & d(2) & "," & d(3) & "," & d(4) & "," & d(5) & "," & d(6) & ")"
        Else
            sql = "UPDATE tour SET pkid=" & d(1) & ",dodepature=" & d(2) & ",doarrival=" & d(3) & " ,BUSNO=" & d(4) & ",totalseats=" & d(5) & ",FARE=" & d(6) & "where tid=" & d(0)
        End If
    
On Error GoTo databaseErrors
    
        Set rstemp = cn.Execute(sql)
              
'        putDataToComboBox
        
        PUTDATATOCMBTID
        
         
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
    
    
    sql = "SELECT * FROM tour WHERE tid='" & t(0) & "'"
   
    Set rstemp = cn.Execute(sql)
    
    If recordCheck(rstemp) = True Then
            
        displayAllRecords t, rstemp
        
            
    'Else
        
    End If
    


End Sub

Private Sub cmdtour_Click(Index As Integer)
     If t(0).Text = "" Then
        MsgBox "Choose Some Tour"
    Else
        showTOURReport t(0).Text
    End If
End Sub

Private Sub Form_Load() 'www.freestudentprojects.com
    
    PUTDATATOCMBTID
    
    putDataToComboBox

GETTOURID

    
End Sub

Private Sub GETTOURID()
  Dim rstemp As New ADODB.Recordset
    
    Set rstemp = cn.Execute("select max(Tid) mTid from tour")
    
   
    If recordCheck(rstemp) = True Then
        If IsNull(rstemp.Fields(0)) = True Then
            t(0).Text = 1
        Else
            t(0).Text = rstemp.Fields(0) + 1
        End If
    Else
            t(0).Text = 1
    End If


End Sub
Private Sub PUTDATATOCMBTID()
    Dim r As New ADODB.Recordset
     Set r = cn.Execute("select tid from tour")
     
      cmbtid.Clear
      
       If recordCheck(r) = True Then
        loadDataToSingleComboBox r, cmbtid
      
    End If
       
End Sub
Private Sub putDataToComboBox()

       Dim rst As New ADODB.Recordset
   
      Set rst = cn.Execute("select BUSNO from BUS")
   
        cmbBusNo.Clear
    
    If recordCheck(rst) = True Then
        
             loadDataToSingleComboBox rst, cmbBusNo
    Else
    
    End If
    Set rst = Nothing
    
End Sub



 
    
