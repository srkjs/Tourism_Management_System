VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} rptenv 
   ClientHeight    =   7500
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7245
   _ExtentX        =   12779
   _ExtentY        =   13229
   FolderFlags     =   1
   TypeLibGuid     =   "{F1DAB7FC-708F-11D6-A69E-000021E5805A}"
   TypeInfoGuid    =   "{F1DAB7FD-708F-11D6-A69E-000021E5805A}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "rptcnn"
      ConnDispId      =   1001
      SourceOfData    =   3
      ConnectionSource=   "Provider=MSDAORA.1;Password=TOURISM;User ID=tourism;Persist Security Info=True"
      Expanded        =   -1  'True
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   3
   BeginProperty Recordset1 
      CommandName     =   "cmdpackage"
      CommDispId      =   1002
      RsDispId        =   1006
      CommandText     =   $"rptenv.dsx":0000
      ActiveConnectionName=   "rptcnn"
      CommandType     =   1
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   5
      BeginProperty Field1 
         Precision       =   2
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "PKID"
         Caption         =   "PKID"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "TNAME"
         Caption         =   "TNAME"
      EndProperty
      BeginProperty Field3 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "FARE"
         Caption         =   "FARE"
      EndProperty
      BeginProperty Field4 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TKM"
         Caption         =   "TKM"
      EndProperty
      BeginProperty Field5 
         Precision       =   3
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "DURATION"
         Caption         =   "DURATION"
      EndProperty
      NumGroups       =   0
      ParamCount      =   1
      BeginProperty P1 
         RealName        =   "Param1"
         UserName        =   "pkid"
         Direction       =   1
         Precision       =   0
         Scale           =   0
         Size            =   0
         DataType        =   131
         HostType        =   3
         Required        =   -1  'True
      EndProperty
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "cmdtourplace"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   $"rptenv.dsx":002A
      ActiveConnectionName=   "rptcnn"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "cmdpackage"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   8
      BeginProperty Field1 
         Precision       =   2
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "PKID"
         Caption         =   "PKID"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "TNAME"
         Caption         =   "TNAME"
      EndProperty
      BeginProperty Field3 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "FARE"
         Caption         =   "FARE"
      EndProperty
      BeginProperty Field4 
         Precision       =   8
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "TKM"
         Caption         =   "TKM"
      EndProperty
      BeginProperty Field5 
         Precision       =   3
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "DURATION"
         Caption         =   "DURATION"
      EndProperty
      BeginProperty Field6 
         Precision       =   4
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "PLID"
         Caption         =   "PLID"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   30
         Scale           =   0
         Type            =   200
         Name            =   "PLACE"
         Caption         =   "PLACE"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "SEASON"
         Caption         =   "SEASON"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "PKID"
         ChildField      =   "PKID"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset3 
      CommandName     =   "cmdTourSite"
      CommDispId      =   -1
      RsDispId        =   -1
      CommandText     =   $"rptenv.dsx":004C
      ActiveConnectionName=   "rptcnn"
      CommandType     =   1
      RelateToParent  =   -1  'True
      ParentCommandName=   "cmdtourplace"
      Expanded        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   3
      BeginProperty Field1 
         Precision       =   4
         Size            =   19
         Scale           =   0
         Type            =   131
         Name            =   "PLID"
         Caption         =   "PLID"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   50
         Scale           =   0
         Type            =   200
         Name            =   "SITE"
         Caption         =   "SITE"
      EndProperty
      BeginProperty Field3 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "SITETYPE"
         Caption         =   "SITETYPE"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   1
      BeginProperty Relation1 
         ParentField     =   "PLID"
         ChildField      =   "PLID"
         ParentType      =   0
         ChildType       =   0
      EndProperty
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "rptenv"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub DataEnvironment_Initialize()

End Sub
