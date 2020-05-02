VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmplace 
   BackColor       =   &H00808080&
   Caption         =   "TOUR PLACES"
   ClientHeight    =   4695
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   ForeColor       =   &H00808080&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MDIChild        =   -1  'True
   ScaleHeight     =   4695
   ScaleWidth      =   5700
   WindowState     =   2  'Maximized
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
      Left            =   9120
      TabIndex        =   21
      Tag             =   "AFD"
      Top             =   6240
      Width           =   1335
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
      Left            =   3840
      TabIndex        =   13
      Top             =   4920
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
      Left            =   2640
      TabIndex        =   12
      Top             =   4920
      Width           =   975
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
      Left            =   1320
      TabIndex        =   11
      Top             =   4920
      Width           =   975
   End
   Begin VB.ComboBox cmbSiteType 
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
      ItemData        =   "frmplace.frx":0000
      Left            =   4920
      List            =   "frmplace.frx":0013
      TabIndex        =   20
      Text            =   "Combo1"
      Top             =   4200
      Width           =   2175
   End
   Begin VB.TextBox ts 
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
      Left            =   2760
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   4200
      Width           =   2055
   End
   Begin VB.TextBox ts 
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
      Left            =   2760
      TabIndex        =   8
      Top             =   3480
      Width           =   5055
   End
   Begin MSFlexGridLib.MSFlexGrid MSF 
      Height          =   1575
      Left            =   1080
      TabIndex        =   19
      Top             =   5520
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2778
      _Version        =   393216
      FixedCols       =   0
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
      Left            =   9360
      TabIndex        =   17
      Tag             =   "AFD"
      Top             =   3960
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
      Left            =   9360
      TabIndex        =   18
      Tag             =   "AFD"
      Top             =   4800
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
      Left            =   9360
      TabIndex        =   16
      Tag             =   "SC"
      Top             =   3000
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
      Left            =   9360
      TabIndex        =   15
      Tag             =   "SC"
      Top             =   2160
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
      Left            =   9360
      TabIndex        =   14
      Tag             =   "AFD"
      Top             =   1200
      Width           =   1215
   End
   Begin VB.ComboBox cmbpid 
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
      Left            =   5880
      TabIndex        =   5
      Top             =   960
      Width           =   2055
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
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   6
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
      Height          =   465
      Index           =   1
      Left            =   2640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1680
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
      Height          =   435
      Index           =   0
      Left            =   2640
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   960
      Width           =   3015
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H0080FFFF&
      Height          =   2415
      Left            =   840
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   7335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H008080FF&
      FillColor       =   &H0000FFFF&
      Height          =   4815
      Left            =   9120
      Shape           =   4  'Rounded Rectangle
      Top             =   960
      Width           =   1695
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00FF8080&
      X1              =   720
      X2              =   8760
      Y1              =   3240
      Y2              =   3240
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Season"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   285
      Left            =   1695
      TabIndex        =   4
      Top             =   2400
      Width           =   720
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Site Type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   285
      Left            =   1455
      TabIndex        =   9
      Top             =   4200
      Width           =   960
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Visiting Site"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   285
      Left            =   1200
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Place Id"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   285
      Left            =   1620
      TabIndex        =   0
      Top             =   960
      Width           =   810
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00808080&
      BackStyle       =   0  'Transparent
      Caption         =   "Visiting Place"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   285
      Left            =   1065
      TabIndex        =   2
      Top             =   1680
      Width           =   1365
   End
End
Attribute VB_Name = "frmplace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim optflag As Integer
Dim c As Integer
Dim mFlag As Boolean
Private Sub cmdadd_Click()
    optflag = 1
End Sub
Private Sub cmbpid_click()
    t(0).Text = cmbpid.Text
    findRecord
End Sub
Private Sub cmbpt_Click()
    t(2).Text = cmbpt.Text
End Sub

Private Sub cmbSiteType_Click()
    ts(1).Text = cmbSiteType.Text
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
            MSF.Clear
            MSF.Rows = 2
            
            getPlaceId
            
            flexGridHeading
            
                      
        Else
            optflag = 2
        End If
    
 
    Else
    
On Error GoTo databaseErrors
    
        If (t(0).Text = "") Then
            MsgBox "Select Some tourplace", vbInformation
        Else
            
            cn.Execute "Delete from tourplace_site where plid='" & t(0).Text & "'"
            
            cn.Execute "delete from tourplace where plid='" & t(0).Text & "'"
      
            MSF.Clear
            MSF.Rows = 2
            flexGridHeading
            putDataToComboBox
            
            ts(0).Text = Clear
            ts(1).Text = Clear
            
            resetAllData t
                   
        End If
    
    End If
    
databaseErrors:
    databaseError


End Sub

Private Sub cmdclose_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdMsfAdd_Click()

    c = c + 1

    MSF.Row = c

    MSF.Col = 0
        MSF.Text = ts(0).Text
    MSF.Col = 1
        MSF.Text = ts(1).Text
        
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
        MSF.Col = 0
            MSF.Text = ts(0).Text
        MSF.Col = 1
            MSF.Text = ts(1).Text
            
        mFlag = False
    End If

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
                If (I = 0) Then
                    d(I) = t(I).Text
                Else
                    d(I) = "'" & t(I).Text & "'"
                End If
            End If
        Next

        If optflag = 1 Then
            sql = "INSERT INTO tourplace VALUES(" & d(0) & "," & d(1) & "," & d(2) & ")"
        Else
            sql = "UPDATE tourplace SET place=" & d(1) & ",season=" & d(2) & " where plid=" & d(0)
        End If
    
On Error GoTo databaseErrors
    
        Set rstemp = cn.Execute(sql)
 
        saveToTourSiteTable
        
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
    
    c = 0
    
End Sub
Private Sub findRecord()

    Dim sql As String
    Dim rstemp As New ADODB.Recordset
    
    sql = "SELECT * FROM tourplace WHERE plid='" & t(0) & "'"
   
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
    
End Sub
Private Sub putDataToComboBox()

    Dim rstemp As New ADODB.Recordset
    
    
    Set rstemp = cn.Execute("select plid from tourplace")
    
    cmbpid.Clear
    
     
    If recordCheck(rstemp) = True Then
        
        loadDataToSingleComboBox rstemp, cmbpid
        
    End If

End Sub
Private Sub flexGridHeading()

    setColumnWidth

    MSF.Row = 0
        MSF.Col = 0
            MSF.Text = "Visitiing Site"
        MSF.Col = 1
            MSF.Text = "Site Type"
            
    MSF.Row = 0

End Sub
Private Sub setColumnWidth()

    Dim I As Integer
    
    For I = 0 To MSF.Cols - 1
    
        MSF.ColWidth(I) = 4000
        
    
    Next
    

End Sub

Private Sub MSF_DblClick()
    If MSF.Row <> 0 And MSF.Row <> MSF.Rows - 1 Then
    
        mFlag = True
    
        If MSF.Text <> "" Then
            MSF.Col = 0
                ts(0).Text = MSF.Text
            MSF.Col = 1
                ts(1).Text = MSF.Text
            
        End If
    End If
End Sub
Private Sub saveToTourSiteTable()

    Dim sql As String
    Dim I As Integer
    Dim d(3) As String
    Dim rstemp As New ADODB.Recordset
    
    Set rstemp = cn.Execute("Delete from tourplace_site where plid=" & t(0).Text)
    
    For I = 1 To MSF.Rows - 2
    
        MSF.Row = I
        
            d(0) = t(0)
            
            MSF.Col = 0
            
            d(1) = "'" & MSF.Text & "'"
            
            MSF.Col = 1
            
            d(2) = "'" & MSF.Text & "'"
    
            sql = "insert into tourplace_site values ( " & d(0) & "," & d(1) & "," & d(2) & ")"
            
            Set rstemp = cn.Execute(sql)
    
    Next
    
    Set rstemp = Nothing

End Sub

Public Sub showDataInFlexGrid()

    Dim sql As String
    Dim I As Integer
    Dim rstemp As New ADODB.Recordset
    
    sql = "SELECT * FROM TOURPLACE_SITE WHERE PLID=" & t(0).Text
    
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
            MSF.Rows = MSF.Rows + 1
            
            rstemp.MoveNext
        Loop
    
    End If


End Sub
Private Sub getPlaceId()
  Dim rstemp As New ADODB.Recordset
    
    Set rstemp = cn.Execute("select max(plid) mplid from tourplace")
    
   
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

Private Sub Picture1_Click()

End Sub

Private Sub Picture2_Click()

End Sub
