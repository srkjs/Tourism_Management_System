VERSION 5.00
Begin VB.Form frmpack 
   BackColor       =   &H00404080&
   Caption         =   "PACKAGE DETAILS"
   ClientHeight    =   6165
   ClientLeft      =   60
   ClientTop       =   360
   ClientWidth     =   6825
   FillColor       =   &H00C0E0FF&
   ForeColor       =   &H00004040&
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6165
   ScaleWidth      =   6825
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbPkid 
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
      Left            =   7080
      TabIndex        =   30
      Text            =   "Combo1"
      Top             =   2040
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
      Height          =   405
      Index           =   5
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   29
      Top             =   3960
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
      Height          =   405
      Index           =   1
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   27
      Top             =   1320
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
      Height          =   405
      Index           =   0
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   26
      Top             =   2040
      Width           =   2655
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
      ItemData        =   "frmpack.frx":0000
      Left            =   4560
      List            =   "frmpack.frx":000D
      TabIndex        =   25
      Text            =   "Combo1"
      Top             =   720
      Width           =   2655
   End
   Begin VB.CommandButton cmdRemovePlace 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   23
      Top             =   6840
      Width           =   615
   End
   Begin VB.CommandButton cmdAppendPlace 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   22
      Top             =   5880
      Width           =   615
   End
   Begin VB.ListBox lstSelPlace 
      Height          =   2400
      Left            =   6720
      TabIndex        =   21
      Top             =   5280
      Width           =   1575
   End
   Begin VB.ListBox lstSelPid 
      Enabled         =   0   'False
      Height          =   2400
      Left            =   5760
      TabIndex        =   17
      Top             =   5280
      Width           =   855
   End
   Begin VB.ListBox lstChoicePlace 
      Height          =   2400
      Left            =   2880
      TabIndex        =   16
      Top             =   5280
      Width           =   1695
   End
   Begin VB.ListBox lstChoicePid 
      Enabled         =   0   'False
      Height          =   2400
      Left            =   1800
      TabIndex        =   15
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton cmdpack 
      Caption         =   "&Package"
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
      TabIndex        =   14
      Tag             =   "AFD"
      Top             =   3360
      Width           =   1095
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
      Left            =   10080
      TabIndex        =   12
      Tag             =   "AFD"
      Top             =   6120
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
      Left            =   10080
      TabIndex        =   10
      Tag             =   "AFD"
      Top             =   4320
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
      Left            =   10080
      TabIndex        =   11
      Tag             =   "AFD"
      Top             =   5280
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
      Left            =   10080
      TabIndex        =   9
      Tag             =   "SC"
      Top             =   3480
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
      Left            =   10080
      TabIndex        =   8
      Tag             =   "SC"
      Top             =   2520
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
      Left            =   10080
      TabIndex        =   7
      Tag             =   "AFD"
      Top             =   1680
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
      Height          =   405
      Index           =   4
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   3360
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
      Height          =   405
      Index           =   2
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   720
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
      Height          =   405
      Index           =   3
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   2760
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFF80&
      Height          =   4095
      Left            =   1800
      Shape           =   4  'Rounded Rectangle
      Top             =   600
      Width           =   6015
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "PACKAGE ID"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   225
      Left            =   2880
      TabIndex        =   28
      Top             =   2160
      Width           =   1140
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "PACKAGE TYPE"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   225
      Left            =   2640
      TabIndex        =   24
      Top             =   840
      Width           =   1410
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Visiting Places"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   5640
      TabIndex        =   20
      Top             =   4920
      Width           =   1305
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Place List"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   3000
      TabIndex        =   19
      Top             =   4920
      Width           =   885
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "Place Id"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   1800
      TabIndex        =   18
      Top             =   4920
      Width           =   735
   End
   Begin VB.Line Line5 
      BorderColor     =   &H00C0C000&
      X1              =   1800
      X2              =   7800
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   "ADDING NEW PACKAGE DETAIL"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   495
      Left            =   1920
      TabIndex        =   13
      Top             =   0
      Width           =   5895
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF80FF&
      Height          =   5415
      Left            =   9960
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   " PACKAGE NAME"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   225
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   " NUMBER OF DAYS "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   225
      Left            =   2280
      TabIndex        =   5
      Top             =   4080
      Width           =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   " FARE "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   225
      Left            =   3360
      TabIndex        =   1
      Top             =   2880
      Width           =   570
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00404080&
      BackStyle       =   0  'Transparent
      Caption         =   " TOTAL KMS "
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   225
      Left            =   2760
      TabIndex        =   3
      Top             =   3480
      Width           =   1200
   End
End
Attribute VB_Name = "frmpack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim optflag As Integer
Private Sub cmdadd_Click()
    optflag = 1
End Sub

Private Sub cmbpkid_Click()
    If optflag = 0 Or optflag = 2 Then
        t(0).Text = cmbpkid.Text
        findRecord
    End If
End Sub

Private Sub cmbpktype_Click()
    t(2).Text = cmbpktype.Text
End Sub
Private Sub loadDataToPackageId()

    Dim rstemp As New ADODB.Recordset
    
    Dim sql As String
    
    sql = "SELECT PKID FROM PACKAGES"
    
    Set rstemp = cn.Execute(sql)
    
    If recordCheck(rstemp) = True Then
        Do Until rstemp.EOF = True
            cmbpkid.AddItem rstemp.Fields(0)
            rstemp.MoveNext
        Loop
    End If


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
            getpkid
        Else
            optflag = 2
        End If
    
 
    Else
    
On Error GoTo databaseErrors
    
        If (t(0).Text = "") Then
            MsgBox "Select Some packages ", vbInformation
        Else
            cn.Execute "DELETE FROM PACKAGE_PLACE WHERE PKID=" & t(0).Text
            cn.Execute "Delete from packages where pkid=" & t(1).Text
                        
            loadDataToPackageId
            resetAllData t
        End If
    
    End If
    
databaseErrors:
    databaseError


End Sub

Private Sub cmdAppendPlace_Click()

    lstSelPid.AddItem lstChoicePid.List(lstChoicePlace.ListIndex)
    lstSelPlace.AddItem lstChoicePlace.List(lstChoicePlace.ListIndex)
    
    lstChoicePid.RemoveItem (lstChoicePlace.ListIndex)
    lstChoicePlace.RemoveItem (lstChoicePlace.ListIndex)


End Sub

Private Sub cmdclose_Click(Index As Integer)
    Unload Me
End Sub

Private Sub cmdpack_Click(Index As Integer)
    If t(1).Text = "" Then
        MsgBox "choose some package ID"
    Else
        showpackReport t(0).Text
    End If
End Sub

Private Sub cmdRemovePlace_Click()
    
    lstChoicePid.AddItem lstSelPid.List(lstSelPlace.ListIndex)
    lstChoicePlace.AddItem lstSelPlace.List(lstSelPlace.ListIndex)
    
    
    lstSelPid.RemoveItem lstSelPlace.ListIndex
    lstSelPlace.RemoveItem lstSelPlace.ListIndex


End Sub

Private Sub cmdSC_Click(Index As Integer)

    Dim rstemp As New ADODB.Recordset
    Dim d(6) As String
    Dim I As Integer
    Dim sql As String
    
    t(2).Text = cmbpktype.Text
    
    If Index = 0 Then
        
         
        For I = 0 To 5
            If (t(I).Text = "") Then
                d(I) = "NULL"
            Else
                If (I = 0 Or I = 3 Or I = 4 Or I = 5) Then
                    d(I) = t(I).Text
                Else
                    d(I) = "'" & t(I).Text & "'"
                End If
            End If
        Next

        If optflag = 1 Then
            sql = "INSERT INTO packages VALUES(" & d(0) & "," & d(1) & "," & d(2) & "," & d(3) & "," & d(4) & "," & d(5) & ")"
        Else
            sql = "UPDATE packages SET  tname=" & d(1) & ",pkname=" & d(2) & ",fare=" & d(3) & ",tkm=" & d(4) & " ,duration=" & d(5) & " where pkid=" & d(0) & ""
        End If
    
On Error GoTo databaseErrors
    
        Set rstemp = cn.Execute(sql)
        
        
        loadDataToPackageId
        
        saveDataToPackagePlace
        
      '  putDataToComboBox
        
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
Private Sub saveDataToPackagePlace()

    Dim sql As String
    Dim rstemp As New ADODB.Recordset
    Dim I As Integer
    Dim d(1) As String
    
    Set rstemp = cn.Execute("Delete from PACKAGE_PLACE WHERE PKID=" & t(0).Text)
    
    For I = 0 To lstSelPid.ListCount - 1
        sql = "insert into package_place values(" & t(0).Text & "," & lstSelPid.List(I) & ")"
        Set rstemp = cn.Execute(sql)
    Next
    
    Set rstemp = Nothing

End Sub
Private Sub findRecord()

    Dim sql As String
    Dim rstemp As New ADODB.Recordset
    
    sql = "SELECT * FROM packages WHERE pkid=" & t(0).Text
    
    Set rstemp = cn.Execute(sql)
    
    If recordCheck(rstemp) = True Then
            
        displayAllRecords t, rstemp
            
        showPlaces
            
    Else
        
    End If
End Sub
Private Sub showPlaces()
    Dim rstemp As New ADODB.Recordset
    Dim sql As String
    Dim I As Integer
    
    sql = "SELECT PLID FROM PACKAGE_PLACE WHERE PKID=" & t(0).Text
    
    Set rstemp = cn.Execute(sql)
    
    loadDataToListBoxes
    
    lstSelPid.Clear
    lstSelPlace.Clear
    
    If recordCheck(rstemp) = True Then
    
        Do Until rstemp.EOF = True
            For I = 0 To lstChoicePid.ListCount - 1
                If rstemp.Fields(0) = lstChoicePid.List(I) Then
                    lstSelPid.AddItem lstChoicePid.List(I)
                    lstSelPlace.AddItem lstChoicePlace.List(I)
                    
                    lstChoicePid.RemoveItem I
                    lstChoicePlace.RemoveItem I
                    Exit For
                End If
            Next
        
            rstemp.MoveNext
        Loop
    
    End If
    
    Set rstemp = Nothing

End Sub
Private Sub Form_Load() 'www.freestudentprojects.com
    loadDataToListBoxes
   
    loadDataToPackageId
   
    'putDataToComboBox
    
End Sub
Private Sub getpkid()
   Dim rstemp As New ADODB.Recordset
    Set rstemp = cn.Execute("SELECT MAX(pkid) FROM packages ")
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
Private Sub loadDataToListBoxes()

    Dim rstemp As New ADODB.Recordset
    
    Set rstemp = cn.Execute("SELECT PLID,PLACE FROM TOURPLACE")
    
    lstChoicePid.Clear
    lstChoicePlace.Clear
    
    If recordCheck(rstemp) = True Then
        
            Do Until rstemp.EOF = True
                lstChoicePid.AddItem rstemp.Fields(0)
                lstChoicePlace.AddItem rstemp.Fields(1)
                rstemp.MoveNext
            Loop
    
    End If

End Sub

