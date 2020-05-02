VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.MDIForm frmMain 
   BackColor       =   &H8000000C&
   Caption         =   "Main Screen Of Tourism"
   ClientHeight    =   3195
   ClientLeft      =   225
   ClientTop       =   1155
   ClientWidth     =   4680
   LinkTopic       =   "MDIForm1"
   LockControls    =   -1  'True
   StartUpPosition =   3  'Windows Default
   Tag             =   "p"
   WindowState     =   2  'Maximized
   Begin VB.PictureBox sbStatusBar 
      Align           =   2  'Align Bottom
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   0
      ScaleHeight     =   210
      ScaleWidth      =   4620
      TabIndex        =   0
      Top             =   2925
      Width           =   4680
   End
   Begin MSComDlg.CommonDialog dlgCommonDialog 
      Left            =   960
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox imlToolbarIcons 
      Align           =   1  'Align Top
      BackColor       =   &H80000005&
      Height          =   480
      Left            =   0
      ScaleHeight     =   420
      ScaleWidth      =   4620
      TabIndex        =   1
      Top             =   0
      Width           =   4680
   End
   Begin VB.Menu mnuOrganization 
      Caption         =   "&Organization"
      Begin VB.Menu optOffice 
         Caption         =   "&Office"
         Shortcut        =   ^O
      End
      Begin VB.Menu optsep1 
         Caption         =   "-"
      End
      Begin VB.Menu optAboutOrg 
         Caption         =   "&About Agency"
         Shortcut        =   ^A
      End
      Begin VB.Menu optSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTourDetails 
      Caption         =   "&Tour-Details"
      Begin VB.Menu optBus 
         Caption         =   "&Bus"
         Shortcut        =   ^B
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu optPlaces 
         Caption         =   "&Places"
         Shortcut        =   ^P
      End
      Begin VB.Menu sep4 
         Caption         =   "-"
      End
      Begin VB.Menu optTourPack 
         Caption         =   "&Tour Packages"
         Shortcut        =   ^T
      End
      Begin VB.Menu sep5 
         Caption         =   "-"
      End
      Begin VB.Menu optpackinfo 
         Caption         =   "&Package Info"
      End
      Begin VB.Menu sep6 
         Caption         =   "-"
      End
      Begin VB.Menu optTourDetail 
         Caption         =   "Tour &Detail"
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu mnuBilling 
      Caption         =   "&Billing"
      Begin VB.Menu optBill 
         Caption         =   "&Bill"
         Shortcut        =   ^L
      End
      Begin VB.Menu sep7 
         Caption         =   "-"
      End
      Begin VB.Menu optReceipt 
         Caption         =   "&Receipt"
         Shortcut        =   ^R
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      WindowList      =   -1  'True
      Begin VB.Menu mnuWindowCascade 
         Caption         =   "&Cascade"
      End
      Begin VB.Menu mnuWindowTileHorizontal 
         Caption         =   "Tile &Horizontal"
      End
      Begin VB.Menu mnuWindowTileVertical 
         Caption         =   "Tile &Vertical"
      End
      Begin VB.Menu mnuWindowArrangeIcons 
         Caption         =   "&Arrange Icons"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
      End
      Begin VB.Menu mnuHelpSearchForHelpOn 
         Caption         =   "&Search For Help On..."
      End
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "&About "
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function OSWinHelp% Lib "user32" Alias "WinHelpA" (ByVal hwnd&, ByVal HelpFile$, ByVal wCommand%, dwData As Any)

Private Sub MDIForm_Load() 'www.freestudentprojects.com
    Me.Left = GetSetting(App.Title, "Settings", "MainLeft", 1000)
    Me.Top = GetSetting(App.Title, "Settings", "MainTop", 1000)
    Me.Width = GetSetting(App.Title, "Settings", "MainWidth", 6500)
    Me.Height = GetSetting(App.Title, "Settings", "MainHeight", 6500)
 
End Sub



Private Sub MDIForm_Unload(Cancel As Integer)
    If Me.WindowState <> vbMinimized Then
        SaveSetting App.Title, "Settings", "MainLeft", Me.Left
        SaveSetting App.Title, "Settings", "MainTop", Me.Top
        SaveSetting App.Title, "Settings", "MainWidth", Me.Width
        SaveSetting App.Title, "Settings", "MainHeight", Me.Height
    End If
    
    destroyAllForms
    
End Sub

Private Sub tbToolBar_ButtonClick(ByVal Button As MSComCtlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            LoadNewDoc
        Case "Open"
            mnuFileOpen_Click
        Case "Save"
            mnuFileSave_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Cut"
            'ToDo: Add 'Cut' button code.
            MsgBox "Add 'Cut' button code."
        Case "Copy"
            'ToDo: Add 'Copy' button code.
            MsgBox "Add 'Copy' button code."
        Case "Paste"
            'ToDo: Add 'Paste' button code.
            MsgBox "Add 'Paste' button code."
        Case "Bold"
            ActiveForm.rtfText.SelBold = Not ActiveForm.rtfText.SelBold
            Button.Value = IIf(ActiveForm.rtfText.SelBold, tbrPressed, tbrUnpressed)
        Case "Italic"
            ActiveForm.rtfText.SelItalic = Not ActiveForm.rtfText.SelItalic
            Button.Value = IIf(ActiveForm.rtfText.SelItalic, tbrPressed, tbrUnpressed)
        Case "Underline"
            ActiveForm.rtfText.SelUnderline = Not ActiveForm.rtfText.SelUnderline
            Button.Value = IIf(ActiveForm.rtfText.SelUnderline, tbrPressed, tbrUnpressed)
        Case "Align Left"
            ActiveForm.rtfText.SelAlignment = rtfLeft
        Case "Center"
            ActiveForm.rtfText.SelAlignment = rtfCenter
        Case "Align Right"
            ActiveForm.rtfText.SelAlignment = rtfRight
    End Select
End Sub

Private Sub mnuFileNew_Click()

End Sub

Private Sub mnuHelpSearchForHelpOn_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 261, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub

Private Sub mnuHelpContents_Click()
    Dim nRet As Integer


    'if there is no helpfile for this project display a message to the user
    'you can set the HelpFile for your application in the
    'Project Properties dialog
    If Len(App.HelpFile) = 0 Then
        MsgBox "Unable to display Help Contents. There is no Help associated with this project.", vbInformation, Me.Caption
    Else
        On Error Resume Next
        nRet = OSWinHelp(Me.hwnd, App.HelpFile, 3, 0)
        If Err Then
            MsgBox Err.Description
        End If
    End If

End Sub


Private Sub mnuWindowArrangeIcons_Click()
    Me.Arrange vbArrangeIcons
End Sub

Private Sub mnuWindowTileVertical_Click()
    Me.Arrange vbTileVertical
End Sub

Private Sub mnuWindowTileHorizontal_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuWindowCascade_Click()
    Me.Arrange vbCascade
End Sub
Private Sub mnuFileExit_Click()
    End
    destroyAllForms
End Sub

Private Sub optAboutOrg_Click()

    'frmBrowser.StartingAddress = "c:\jjj.htm"

    'frmBrowser.Show

End Sub

Private Sub optBill_Click()
    frmbill.Show
    frmbill.ZOrder 0
End Sub

Private Sub optBus_Click()
    frmBus.Show
    frmBus.ZOrder 0
End Sub

Private Sub optcustomer_Click()
    rpt.Show
End Sub

Private Sub optOffice_Click()
    frmBranch.Show
    frmBranch.ZOrder 0
End Sub

Private Sub optplace_Click()
    'place.Show
End Sub

Private Sub optpackinfo_Click()
frmpack.Show
frmpack.ZOrder 0

End Sub

Private Sub optPlaces_Click()
    frmplace.Show
    frmplace.ZOrder 0
End Sub

Private Sub optReceipt_Click()
    frmrep.Show
    frmrep.ZOrder 0
End Sub

Private Sub optTourDetail_Click()
    frmtour.Show
    frmtour.ZOrder 0
End Sub

Private Sub optTourPack_Click()
    frmUltraPackage.Show
    frmUltraPackage.ZOrder 0
End Sub
