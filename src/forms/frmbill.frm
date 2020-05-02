VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmbill 
   BackColor       =   &H00C0C0FF&
   Caption         =   "BILL FORMAT"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9480
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   9480
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdBalance 
      Caption         =   "Balance"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   42
      Top             =   2880
      Width           =   1215
   End
   Begin VB.ComboBox cmbBillNumber 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   41
      Text            =   "Combo1"
      Top             =   240
      Width           =   1935
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
      Left            =   10320
      TabIndex        =   40
      Tag             =   "AFD"
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdBill 
      Caption         =   "Bill"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8040
      TabIndex        =   39
      Top             =   2880
      Width           =   1095
   End
   Begin VB.TextBox tBbt 
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
      Left            =   1560
      TabIndex        =   37
      Top             =   4560
      Width           =   3615
   End
   Begin VB.CommandButton cmdMsfAdd 
      Caption         =   "&Add"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   36
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdMsfModify 
      Caption         =   "&Modify"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2040
      TabIndex        =   35
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdMsfDelete 
      Caption         =   "&Delete"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3240
      TabIndex        =   34
      Top             =   5280
      Width           =   975
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
      Left            =   10320
      TabIndex        =   33
      Tag             =   "AFD"
      Top             =   1800
      Width           =   1095
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
      Left            =   10320
      TabIndex        =   32
      Tag             =   "SC"
      Top             =   2640
      Width           =   1095
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
      Left            =   10320
      TabIndex        =   31
      Tag             =   "SC"
      Top             =   3480
      Width           =   1095
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
      Left            =   10320
      TabIndex        =   30
      Tag             =   "AFD"
      Top             =   5160
      Width           =   1095
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
      Left            =   10320
      TabIndex        =   29
      Tag             =   "AFD"
      Top             =   4320
      Width           =   1095
   End
   Begin VB.ComboBox cmbbranchname 
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
      Left            =   5640
      TabIndex        =   28
      Text            =   "COMBO"
      Top             =   240
      Width           =   1932
   End
   Begin MSFlexGridLib.MSFlexGrid msf 
      Height          =   1455
      Left            =   360
      TabIndex        =   27
      Top             =   5880
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   2566
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      GridLines       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox tBbt 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   6600
      TabIndex        =   21
      Top             =   4530
      Width           =   975
   End
   Begin VB.TextBox tBbt 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Index           =   2
      Left            =   7920
      TabIndex        =   25
      Top             =   3840
      Width           =   855
   End
   Begin VB.TextBox tBbt 
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
      Left            =   6120
      TabIndex        =   23
      Top             =   3840
      Width           =   975
   End
   Begin VB.TextBox tBbt 
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
      Index           =   0
      Left            =   1560
      TabIndex        =   19
      Top             =   3840
      Width           =   3615
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
      Left            =   3720
      TabIndex        =   26
      Text            =   "COMBO"
      Top             =   1200
      Width           =   1455
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
      Height          =   372
      Index           =   8
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3000
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
      Height          =   372
      Index           =   7
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   2400
      Width           =   3975
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
      Height          =   405
      Index           =   6
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   1800
      Width           =   3975
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
      Height          =   372
      Index           =   5
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   2160
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
      Height          =   372
      Index           =   4
      Left            =   7680
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   1560
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
      Height          =   372
      Index           =   3
      Left            =   9240
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   600
      Width           =   2055
   End
   Begin VB.TextBox t 
      Height          =   372
      Index           =   2
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   1932
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
      Height          =   372
      Index           =   1
      Left            =   5640
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   600
      Width           =   1932
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
      Height          =   405
      Index           =   0
      Left            =   1680
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   1932
   End
   Begin VB.Shape Shape1 
      Height          =   4935
      Left            =   10200
      Shape           =   4  'Rounded Rectangle
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Seat No"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5520
      TabIndex        =   38
      Top             =   4560
      Width           =   765
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      X1              =   120
      X2              =   9960
      Y1              =   3600
      Y2              =   3600
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " Name "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   600
      TabIndex        =   18
      Top             =   3840
      Width           =   690
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " Relation "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   4560
      Width           =   945
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " Age "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   7320
      TabIndex        =   24
      Top             =   3840
      Width           =   525
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " Sex "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   5520
      TabIndex        =   22
      Top             =   3840
      Width           =   480
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   300
      TabIndex        =   8
      Top             =   3000
      Width           =   1185
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " ADDRESS2 "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   360
      TabIndex        =   6
      Top             =   2400
      Width           =   1080
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " ADDRESS1 "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   360
      TabIndex        =   4
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " BALANCE "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   6480
      TabIndex        =   14
      Top             =   2160
      Width           =   990
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " TOTAL "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   6720
      TabIndex        =   12
      Top             =   1680
      Width           =   735
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " BILL DATE "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   8040
      TabIndex        =   16
      Top             =   600
      Width           =   1065
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
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
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3960
      TabIndex        =   10
      Top             =   600
      Width           =   1485
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   " BILL NUMBER "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   60
      TabIndex        =   0
      Top             =   600
      Width           =   1365
   End
End
Attribute VB_Name = "frmbill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim optflag As Integer
Dim c As Integer
Dim mFlag As Boolean
Dim TOT As Single
Private Sub cmdadd_Click()
    optflag = 1
End Sub

Private Sub cmbBillNumber_Click()
    t(0).Text = cmbBillNumber.Text
    findRecord
End Sub

Private Sub cmbtid_click()
    t(2).Text = cmbtid.Text
    'findRecord
End Sub
Private Sub cmbBRANCHNAME_Click()
    t(1).Text = cmbbranchname.Text
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
            
            getBillNumber
            
            MSF.Clear
            MSF.Rows = 2
            
                        
            flexGridHeading
            
                      
        Else
            optflag = 2
        End If
    
 
    Else
    
On Error GoTo databaseErrors
    
        If (t(0).Text = "") Then
            MsgBox "Select Some BILLNUMBER", vbInformation
        Else
            
            cn.Execute "Delete from receipt where billnum=" & t(0).Text
            
            cn.Execute "delete from booking where billnum=" & t(0).Text
      
            cn.Execute "delete from billmaster where billnum=" & t(0).Text
            
            putDataToBillComboBox
            
            MSF.Clear
            
            MSF.Rows = 2
            
            flexGridHeading
            
            putDataToComboBox
            
            tBbt(0).Text = Clear
            tBbt(1).Text = Clear
            tBbt(2).Text = Clear
            tBbt(3).Text = Clear
            tBbt(4).Text = Clear
            
            resetAllData t
                   
        End If
    
    End If
    
databaseErrors:
    databaseError


End Sub

Private Sub cmdBalance_Click()


    showBalanceReport
    

End Sub

Private Sub cmdBill_Click()
    If t(0).Text = "" Then
        MsgBox "choose some billnum"
    Else
        showBillReport t(0).Text
    End If
End Sub

 
Private Sub cmdclose_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdMsfAdd_Click()

    c = c + 1

    MSF.Row = c

    MSF.Col = 0
        MSF.Text = c
    MSF.Col = 1
        MSF.Text = tBbt(0).Text
     MSF.Col = 2
        MSF.Text = tBbt(1).Text
    MSF.Col = 3
        MSF.Text = tBbt(2).Text
         MSF.Col = 4
        MSF.Text = tBbt(3).Text
         MSF.Col = 5
        MSF.Text = tBbt(4).Text
          
          MSF.Rows = MSF.Rows + 1

End Sub
Private Sub cmdMsfDelete_Click()
    If MSF.Row <> 0 And MSF.Row <> MSF.Rows - 1 Then
        If MSF.Row = 1 And MSF.Rows = 2 Then
            MSF.Col = 0
                MSF.Text = ""
            MSF.Col = 1
                MSF.Text = ""
        Else
            MSF.RemoveItem MSF.Row
            c = c - 1
        End If
    End If

End Sub
Private Sub cmdMsfModify_Click()

    If mFlag = True Then
        MSF.Col = 1
            MSF.Text = tBbt(0).Text
        MSF.Col = 2
            MSF.Text = tBbt(1).Text
        MSF.Col = 3
            MSF.Text = tBbt(2).Text
        MSF.Col = 4
            MSF.Text = tBbt(3).Text
              MSF.Col = 5
            MSF.Text = tBbt(4).Text
        mFlag = False
    End If

End Sub
Private Sub cmdSC_Click(Index As Integer)

    Dim rstemp As New ADODB.Recordset
    Dim d(9) As String
    Dim I As Integer
    Dim sql As String
    
    getTotalAmount
    
    If Index = 0 Then
           
        For I = 0 To 8
            If (t(I).Text = "") Then
                d(I) = "NULL"
            Else
                If (I = 0 Or I = 4 Or I = 5 Or I = 8) Then
                    d(I) = t(I).Text
                Else
                    d(I) = "'" & t(I).Text & "'"
                End If
            End If
        Next

        If optflag = 1 Then
            sql = "INSERT INTO billmaster VALUES(" & d(0) & "," & d(1) & "," & d(2) & "," & d(3) & "," & d(4) & "," & d(5) & "," & d(6) & "," & d(7) & "," & d(8) & ")"
        Else
            sql = "UPDATE billmaster SET pname=" & d(1) & ",tid=" & d(2) & ",dob=" & d(3) & ",total=" & d(4) & ",balances=" & d(5) & ",address1=" & d(6) & " ,address2=" & d(7) & ",telephone=" & d(8) & " where billnum=" & d(0)
        End If
    
On Error GoTo databaseErrors
    
        Set rstemp = cn.Execute(sql)
 
        saveTobookingTable
        
        putDataToComboBox
        
     Else
        
        resetAllData t
    
    End If
    
    
databaseErrors:
    
'    databaseError
    
    lockAllTextBoxes t
    
    setSomeFunctionButtons cmdAFD, "AFD"
    
    resetSomeFunctionButtons cmdSC, "SC"
    
    
    Set rstemp = Nothing
    
    c = 0
End Sub
Private Sub getTotalAmount()

    Dim c As Integer
    Dim sql As String
    Dim rstemp As New ADODB.Recordset
    
    
    For c = 1 To MSF.Rows - 1
    
    Next
    
    Set rstemp = cn.Execute("SELECT FARE FROM TOUR WHERE TID='" & t(2).Text & "'")
    
    If recordCheck(rstemp) = True Then
    
        t(4).Text = c * rstemp.Fields(0)
        
        t(5).Text = t(4).Text
    
    End If

End Sub

Private Sub findRecord()

    Dim sql As String
    Dim rstemp As New ADODB.Recordset
    
    sql = "SELECT * FROM billmaster WHERE billnum='" & t(0) & "'"
   
    Set rstemp = cn.Execute(sql)
    
    If recordCheck(rstemp) = True Then
            
        displayAllRecords t, rstemp
        
        showDataInFlexGrid
        
    Else
        
    End If
End Sub


Private Sub Form_Load() 'www.freestudentprojects.com
    
    flexGridHeading
    
    putDataToComboBox
    
    
    putDataToBillComboBox
    
End Sub
Private Sub putDataToBillComboBox()

    Dim sql As String
    Dim rstemp As New ADODB.Recordset
    
    sql = "SELECT BILLNUM FROM BILLMASTER"
    
    Set rstemp = cn.Execute(sql)
    
    cmbBillNumber.Clear
    
    If recordCheck(rstemp) = True Then
    
        loadDataToSingleComboBox rstemp, cmbBillNumber
    
    End If
    Set rstemp = Nothing
    
   End Sub

Private Sub putDataToComboBox()

    Dim rstemp As New ADODB.Recordset
    Dim rst As New ADODB.Recordset
    
    Set rstemp = cn.Execute("select tid from tour")
    
    Set rst = cn.Execute("select pname from branch")
    
    cmbtid.Clear
    
    cmbbranchname.Clear
     
    If recordCheck(rstemp) = True And recordCheck(rst) = True Then
        
        loadDataToSingleComboBox rstemp, cmbtid
        
        loadDataToSingleComboBox rst, cmbbranchname
    End If
    
End Sub
Private Sub flexGridHeading()

    setColumnWidth

    MSF.Row = 0
        MSF.Col = 0
            MSF.Text = "Sl no."
        MSF.Col = 1
            MSF.Text = "Name"
        MSF.Col = 2
            MSF.Text = "Sex"
        MSF.Col = 3
            MSF.Text = "Age"
                MSF.Col = 4
            MSF.Text = "Seat no."
             MSF.Col = 5
            MSF.Text = "Relation"
    MSF.Row = 0

End Sub
Private Sub setColumnWidth()

    Dim I As Integer
    
    For I = 0 To MSF.Cols - 1
    
        MSF.ColWidth(I) = 1500
        
    Next

End Sub

Private Sub MSF_DblClick()
    If MSF.Row <> 0 And MSF.Row <> MSF.Rows - 1 Then
    
        mFlag = True
    
        If MSF.Text <> "" Then
            MSF.Col = 1
                tBbt(0).Text = MSF.Text
            MSF.Col = 2
                tBbt(1).Text = MSF.Text
            MSF.Col = 3
                tBbt(2).Text = MSF.Text
            MSF.Col = 4
                tBbt(3).Text = MSF.Text
                 MSF.Col = 5
                tBbt(4).Text = MSF.Text
              
        End If
    End If
End Sub
Private Sub saveTobookingTable()

    Dim sql As String
    Dim I As Integer
    Dim d(7) As String
    Dim rstemp As New ADODB.Recordset
  '  Dim TOT As Double
  '  Dim B As Integer
    
   ' Dim rsAmount As New ADODB.Recordset
    
  '  sql = "SELECT FARE FROM TOUR WHERE TID=" & t(2).Text
    
 '   TOT = t(4).Text
    
    'B = B - TOT
   ' Set rsAmount = cn.Execute("UPDATE BILLMASTER SET TOTAL=" & d(4) & " WHERE BILLNUM=" & d(0))
               
    Set rstemp = cn.Execute("Delete from booking where billnum=" & t(0).Text)
       
    For I = 1 To MSF.Rows - 2
    
        MSF.Row = I
        
            d(0) = t(0)
            
            MSF.Col = 0
            
            d(1) = MSF.Text
            
            MSF.Col = 1
            
            d(2) = "'" & MSF.Text & "'"
    
            MSF.Col = 2
            
            d(3) = "'" & MSF.Text & "'"
            
            MSF.Col = 3
            
            d(4) = MSF.Text
            
               MSF.Col = 4
            
            d(5) = MSF.Text
               MSF.Col = 5
            
            d(6) = "'" & MSF.Text & "'"
              
            
            sql = "insert into bOOKING values ( " & d(0) & "," & d(1) & "," & d(2) & "," & d(3) & "," & d(4) & "," & d(5) & "," & d(6) & ")"
            
            Set rstemp = cn.Execute(sql)
    
    Next
    
    Set rstemp = Nothing

End Sub
Public Sub showDataInFlexGrid()

    Dim sql As String
    Dim I As Integer
    Dim rstemp As New ADODB.Recordset
    
    sql = "SELECT * FROM BOOKING WHERE billnum=" & t(0).Text
    
    Set rstemp = cn.Execute(sql)
    
    MSF.Rows = 2
    
    If recordCheck(rstemp) = True Then
    
        Do Until rstemp.EOF = True
            I = I + 1
            MSF.Row = I
                MSF.Col = 0
                    MSF.Text = rstemp.Fields(1)
                MSF.Col = 1
                    MSF.Text = rstemp.Fields(2)
                MSF.Col = 2
                    MSF.Text = rstemp.Fields(3)
                MSF.Col = 3
                    MSF.Text = IIf(IsNull(rstemp.Fields(4)), "", rstemp.Fields(4))
                MSF.Col = 4
                    MSF.Text = IIf(IsNull(rstemp.Fields(5)), "", rstemp.Fields(5))
                MSF.Col = 5
                    MSF.Text = rstemp.Fields(6)
    
            
            MSF.Rows = MSF.Rows + 1
            
            rstemp.MoveNext
        Loop
    
    End If


End Sub
Private Sub getBillNumber()
    Dim rstemp As New ADODB.Recordset
    Set rstemp = cn.Execute("SELECT MAX(BILLNUM) FROM BILLMASTER  ")
    If recordCheck(rstemp) = True Then
        If (IsNull(rstemp.Fields(0)) = True) Then
            t(0).Text = 1
        Else
            t(0).Text = rstemp.Fields(0) + 1
        End If
    Else
        t(0).Text = 0
    End If
    Set rstemp = Nothing
End Sub

