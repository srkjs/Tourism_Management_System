VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   7500
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.Frame fraMainFrame 
      BackColor       =   &H0080C0FF&
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7500
      Begin VB.Timer Timer1 
         Interval        =   7600
         Left            =   6840
         Top             =   4080
      End
      Begin VB.Label lblCopyright 
         BackColor       =   &H0080C0FF&
         Caption         =   "Copyright    All Rights Reserved"
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Tag             =   "Copyright"
         Top             =   3480
         Width           =   2415
      End
      Begin VB.Label lblCompany 
         BackColor       =   &H0080C0FF&
         Caption         =   "Company "
         Height          =   255
         Left            =   3960
         TabIndex        =   9
         Tag             =   "Company"
         Top             =   3840
         Width           =   2415
      End
      Begin VB.Label lblWarning 
         BackStyle       =   0  'Transparent
         Caption         =   "Save Earth"
         Height          =   195
         Left            =   360
         TabIndex        =   8
         Tag             =   "Warning"
         Top             =   4200
         Width           =   6855
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Version"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4560
         TabIndex        =   7
         Tag             =   "Version"
         Top             =   2880
         Width           =   930
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Platform"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4320
         TabIndex        =   6
         Tag             =   "Platform"
         Top             =   2400
         Width           =   1140
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   "Surila Travel Agency"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         TabIndex        =   5
         Tag             =   "CompanyProduct"
         Top             =   765
         Width           =   3630
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         BackColor       =   &H0080C0FF&
         Caption         =   """Tourism"""
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   27.75
            Charset         =   255
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   3240
         TabIndex        =   4
         Tag             =   "Product"
         Top             =   1320
         Width           =   2640
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         Caption         =   "Jawaharlal Nehru National College Of Engineering"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   270
         TabIndex        =   3
         Tag             =   "LicenseTo"
         Top             =   300
         Width           =   6855
      End
      Begin VB.Label Label1 
         BackColor       =   &H0080C0FF&
         Caption         =   "Windows"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   2
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "95 &&Above"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5760
         TabIndex        =   1
         Top             =   3000
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load() 'www.freestudentprojects.com
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblProductName.Caption = App.Title
End Sub

