VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form frmUltraPackage 
   Caption         =   "Vivek Travels - Package Info"
   ClientHeight    =   5880
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   9075
   StartUpPosition =   3  'Windows Default
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
      Left            =   7080
      TabIndex        =   26
      Tag             =   "AFD"
      Top             =   480
      Width           =   1095
   End
   Begin VB.TextBox txtFare 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   7560
      TabIndex        =   22
      Top             =   1560
      Width           =   735
   End
   Begin VB.TextBox txtFare 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   4320
      TabIndex        =   21
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox txtFare 
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1920
      TabIndex        =   18
      Top             =   1560
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid mss 
      Height          =   1575
      Left            =   1440
      TabIndex        =   17
      Top             =   3960
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2778
      _Version        =   393216
      FixedCols       =   0
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSFlexGridLib.MSFlexGrid msp 
      Height          =   1575
      Left            =   1440
      TabIndex        =   5
      Top             =   2280
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   2778
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ForeColor       =   8421376
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.ComboBox cmbpkid 
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
      Left            =   2400
      TabIndex        =   2
      Text            =   "Combo2"
      Top             =   960
      Width           =   2775
   End
   Begin VB.ComboBox cmbpktype 
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
      ItemData        =   "frmUltraPackage.frx":0000
      Left            =   1920
      List            =   "frmUltraPackage.frx":000D
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   360
      Width           =   2775
   End
   Begin VB.ComboBox cmbPkiiid 
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
      Left            =   1920
      TabIndex        =   4
      Text            =   "Combo2"
      Top             =   960
      Width           =   2775
   End
   Begin VB.Image Image24 
      Height          =   1725
      Left            =   8400
      Picture         =   "frmUltraPackage.frx":0043
      Top             =   5040
      Width           =   1725
   End
   Begin VB.Image Image23 
      Height          =   1725
      Left            =   3360
      Picture         =   "frmUltraPackage.frx":108F
      Top             =   5040
      Width           =   1725
   End
   Begin VB.Image Image22 
      Height          =   1725
      Left            =   6720
      Picture         =   "frmUltraPackage.frx":20DB
      Top             =   5040
      Width           =   1725
   End
   Begin VB.Image Image18 
      Height          =   1725
      Left            =   5040
      Picture         =   "frmUltraPackage.frx":3127
      Top             =   5040
      Width           =   1725
   End
   Begin VB.Image Image17 
      Height          =   1725
      Left            =   1680
      Picture         =   "frmUltraPackage.frx":4173
      Top             =   5040
      Width           =   1725
   End
   Begin VB.Image Image16 
      Height          =   1725
      Left            =   0
      Picture         =   "frmUltraPackage.frx":51BF
      Top             =   5040
      Width           =   1725
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   5640
      TabIndex        =   25
      Top             =   1560
      Width           =   495
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Duration"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3360
      TabIndex        =   24
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Rs."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1560
      TabIndex        =   23
      Top             =   1560
      Width           =   375
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Km."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6600
      TabIndex        =   20
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Fare"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   480
      TabIndex        =   19
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
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
      Left            =   1080
      TabIndex        =   16
      Top             =   4800
      Width           =   120
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
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
      Left            =   1080
      TabIndex        =   15
      Top             =   4680
      Width           =   105
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T"
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
      Left            =   1080
      TabIndex        =   14
      Top             =   4440
      Width           =   120
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "I"
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
      Left            =   1080
      TabIndex        =   13
      Top             =   4200
      Width           =   60
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
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
      Left            =   1080
      TabIndex        =   12
      Top             =   3960
      Width           =   120
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "S"
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
      Left            =   1080
      TabIndex        =   11
      Top             =   3480
      Width           =   120
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E"
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
      Left            =   1080
      TabIndex        =   10
      Top             =   3240
      Width           =   105
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
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
      Left            =   1080
      TabIndex        =   9
      Top             =   3000
      Width           =   135
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "A"
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
      Left            =   1080
      TabIndex        =   8
      Top             =   2760
      Width           =   135
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "L"
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
      Left            =   1080
      TabIndex        =   7
      Top             =   2520
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "P"
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
      Left            =   1080
      TabIndex        =   6
      Top             =   2280
      Width           =   120
   End
   Begin VB.Image Image21 
      Height          =   1725
      Left            =   5040
      Picture         =   "frmUltraPackage.frx":620B
      Top             =   3360
      Width           =   1725
   End
   Begin VB.Image Image20 
      Height          =   1725
      Left            =   6720
      Picture         =   "frmUltraPackage.frx":7257
      Top             =   3360
      Width           =   1725
   End
   Begin VB.Image Image19 
      Height          =   1725
      Left            =   8400
      Picture         =   "frmUltraPackage.frx":82A3
      Top             =   3360
      Width           =   1725
   End
   Begin VB.Image Image15 
      Height          =   1725
      Left            =   5040
      Picture         =   "frmUltraPackage.frx":92EF
      Top             =   0
      Width           =   1725
   End
   Begin VB.Image Image14 
      Height          =   1725
      Left            =   6720
      Picture         =   "frmUltraPackage.frx":A33B
      Top             =   0
      Width           =   1725
   End
   Begin VB.Image Image13 
      Height          =   1725
      Left            =   8400
      Picture         =   "frmUltraPackage.frx":B387
      Top             =   0
      Width           =   1725
   End
   Begin VB.Image Image12 
      Height          =   1725
      Left            =   5040
      Picture         =   "frmUltraPackage.frx":C3D3
      Top             =   1680
      Width           =   1725
   End
   Begin VB.Image Image11 
      Height          =   1725
      Left            =   6720
      Picture         =   "frmUltraPackage.frx":D41F
      Top             =   1680
      Width           =   1725
   End
   Begin VB.Image Image10 
      Height          =   1725
      Left            =   8400
      Picture         =   "frmUltraPackage.frx":E46B
      Top             =   1680
      Width           =   1725
   End
   Begin VB.Image Image9 
      Height          =   1725
      Left            =   3360
      Picture         =   "frmUltraPackage.frx":F4B7
      Top             =   3360
      Width           =   1725
   End
   Begin VB.Image Image8 
      Height          =   1725
      Left            =   1680
      Picture         =   "frmUltraPackage.frx":10503
      Top             =   3360
      Width           =   1725
   End
   Begin VB.Image Image7 
      Height          =   1725
      Left            =   0
      Picture         =   "frmUltraPackage.frx":1154F
      Top             =   3360
      Width           =   1725
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Packge Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   240
      TabIndex        =   3
      Top             =   960
      Width           =   1395
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Package Type"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   1695
   End
   Begin VB.Image Image6 
      Height          =   1725
      Left            =   3360
      Picture         =   "frmUltraPackage.frx":1259B
      Top             =   1680
      Width           =   1725
   End
   Begin VB.Image Image5 
      Height          =   1725
      Left            =   1680
      Picture         =   "frmUltraPackage.frx":135E7
      Top             =   1680
      Width           =   1725
   End
   Begin VB.Image Image4 
      Height          =   1725
      Left            =   0
      Picture         =   "frmUltraPackage.frx":14633
      Top             =   1680
      Width           =   1725
   End
   Begin VB.Image Image3 
      Height          =   1725
      Left            =   3360
      Picture         =   "frmUltraPackage.frx":1567F
      Top             =   0
      Width           =   1725
   End
   Begin VB.Image Image2 
      Height          =   1725
      Left            =   1680
      Picture         =   "frmUltraPackage.frx":166CB
      Top             =   0
      Width           =   1725
   End
   Begin VB.Image Image1 
      Height          =   1725
      Left            =   0
      Picture         =   "frmUltraPackage.frx":17717
      Top             =   0
      Width           =   1725
   End
End
Attribute VB_Name = "frmUltraPackage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbpkid_Click()
    cmbPkiiid.ListIndex = cmbpkid.ListIndex
    
    mss.Clear
    mss.Rows = 2
    
    msp.Clear
    msp.Rows = 2
    
    flexGridHeading1
    flexGridHeading2
    
    getPackagePlaces
    
    getPackageDetails
 
 
End Sub
Private Sub getPackageDetails()
    Dim sql As String
    Dim rstemp As New ADODB.Recordset
    Dim I As Integer
    
    sql = "SELECT FARE,DURATION,TKM FROM PACKAGES WHERE PKID=" & cmbPkiiid.Text
    
    Set rstemp = cn.Execute(sql)
    
    If recordCheck(rstemp) = True Then
        For I = 0 To rstemp.Fields.Count - 1
            If IsNull(rstemp.Fields(I)) = True Then
                txtFare(I).Text = ""
            Else
                txtFare(I).Text = rstemp.Fields(I)
            End If
        Next
    End If
    
    Set rstemp = Nothing
End Sub

Private Sub getPackagePlaces()

    Dim rstemp As New ADODB.Recordset
    Dim sql As String
    
    sql = "SELECT PLID,PLACE,SEASON from PACK_PLACE WHERE PKID=" & cmbPkiiid.Text
    
    
    Set rstemp = cn.Execute(sql)
    
    If recordCheck(rstemp) = True Then
        
        loadDataToPlaceFlexGrid rstemp
    
    Else
    
    
    End If


End Sub
Private Sub loadDataToPlaceFlexGrid(r As ADODB.Recordset)
    Dim c As Integer
    
    msp.Clear
    
    flexGridHeading1
    
    msp.Rows = 2
    
    Do Until r.EOF = True
        c = c + 1
        msp.Row = c
            msp.Col = 0
                msp.Text = IIf(IsNull(r.Fields(0)) = True, "", r.Fields(0))
            msp.Col = 1
                msp.Text = IIf(IsNull(r.Fields(1)) = False, r.Fields(1), "")
            msp.Col = 2
                msp.Text = IIf(IsNull(r.Fields(2)) = False, r.Fields(2), "")
        msp.Rows = msp.Rows + 1
        r.MoveNext
    Loop
End Sub


Private Sub cmbpktype_Click()
    Dim sql As String
    Dim rstemp As New ADODB.Recordset
    
    sql = "SELECT tname,pkid FROM packages where pkname='" & cmbpktype.Text & "'"
  
    Set rstemp = cn.Execute(sql)
    
    If recordCheck(rstemp) = True Then
 
        
        loadDataToDoubleComboBox rstemp, cmbpkid, cmbPkiiid
   
    Else
        
    End If
    
    mss.Clear
    mss.Rows = 2
    
    msp.Clear
    msp.Rows = 2
    flexGridHeading1
    flexGridHeading2
    
    Set rstemp = Nothing
End Sub

Private Sub lstPlaces_Click()
    lstPlaceId.ListIndex = lstPlaces.ListIndex
End Sub


Private Sub flexGridHeading1()

    setColumnWidth

    msp.Row = 0
        msp.Col = 0
            msp.Text = "PlaceID"
        msp.Col = 1
            msp.Text = "Visting Places"
        msp.Col = 2
            msp.Text = "Seasons"
            
    msp.Row = 0

End Sub
Private Sub setColumnWidth()

    Dim I As Integer
    
    For I = 0 To msp.Cols - 1
    
        msp.ColWidth(0) = 850
        msp.ColWidth(I) = 2000
        
        
    
    Next
    

End Sub
Private Sub flexGridHeading2()

    setColumnWidth1

    mss.Row = 0
        mss.Col = 0
            mss.Text = "Visitiing Site"
        mss.Col = 1
            mss.Text = "Site Type"
            
    mss.Row = 0

End Sub
Private Sub setColumnWidth1()

    Dim I As Integer
    
    For I = 0 To mss.Cols - 1
    
        mss.ColWidth(I) = 2000
        
    
    Next
    

End Sub

Private Sub cmdpack_Click(Index As Integer)
 If cmbpkid.Text = "" Then
        MsgBox "choose some package ID"
    Else
        showpackReport cmbpkid.Text
    End If
End Sub
Private Sub Form_Load() 'www.freestudentprojects.com

    flexGridHeading1

    flexGridHeading2

End Sub

Private Sub msp_Click()
    If msp.Row <> 0 And msp.Rows - 1 <> 0 Then
        msp.Col = 0
        loadDataToMssFlexGrid msp.Text
    End If

End Sub

Private Sub loadDataToMssFlexGrid(s As String)

    Dim c As Integer
    Dim r As New ADODB.Recordset
    
    Set r = cn.Execute("SELECT * FROM TOURPLACE_SITE WHERE PLID=" & s)
    
    
    mss.Clear
    
    flexGridHeading2
    
    mss.Rows = 2
    
    If recordCheck(r) = True Then
    
        Do Until r.EOF = True
            c = c + 1
            mss.Row = c
                mss.Col = 0
                    mss.Text = IIf(IsNull(r.Fields(1)) = True, "", r.Fields(1))
                mss.Col = 1
                    mss.Text = IIf(IsNull(r.Fields(2)) = False, r.Fields(2), "")
                
            mss.Rows = mss.Rows + 1
            
            r.MoveNext
        Loop
        
    End If
    Set rstemp = Nothing
End Sub
