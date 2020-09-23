VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDeathRecords 
   BackColor       =   &H006B0500&
   BorderStyle     =   0  'None
   Caption         =   "Death Records"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   11235
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Menu"
      Height          =   375
      Left            =   9600
      TabIndex        =   21
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox txtCommentsLink 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      DataField       =   "Comments"
      DataSource      =   "dtaDeathlink"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6840
      TabIndex        =   14
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox txtAgeLink 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      DataField       =   "Age of Death"
      DataSource      =   "dtaDeathlink"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6840
      TabIndex        =   13
      Top             =   240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox txtSurenameLink 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      DataField       =   "Surename"
      DataSource      =   "dtaDeathlink"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1680
      TabIndex        =   12
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.TextBox txtForenamelink 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      DataField       =   "Forename"
      DataSource      =   "dtaDeathlink"
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1680
      TabIndex        =   11
      Top             =   240
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H006B0500&
      Caption         =   "Add New"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   10935
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   5520
         TabIndex        =   19
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtComments 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtAge 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1800
         TabIndex        =   9
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtSurename 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Top             =   840
         Width           =   3495
      End
      Begin VB.TextBox txtForename 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1800
         TabIndex        =   7
         Top             =   480
         Width           =   3495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Comments:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Age At Death:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   5
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Surename:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Forename:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   3
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H006B0500&
      Caption         =   "Death Records"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   3120
      Width           =   10935
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Bindings        =   "frmDeathRecords.frx":0000
         Height          =   2055
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   3625
         _Version        =   393216
         BackColorBkg    =   7013632
         WordWrap        =   -1  'True
         Redraw          =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   ""
      End
   End
   Begin VB.Data dtaDeathlink 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Avalon\My Documents\Higher National Diploma\College\Project\The Genesis Project\KJV.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   315
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "DeathRecords"
      Top             =   2880
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "- Death Records"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Comments:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5520
      TabIndex        =   18
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Age At Death:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   5400
      TabIndex        =   17
      Top             =   240
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Surename:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   16
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Forename:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   360
      TabIndex        =   15
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmDeathRecords"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    Unload Me
    ShapedForm.Show
End Sub

Private Sub cmdSave_Click()


If txtForename.Text = "" Or txtSurename.Text = "" Or txtAge.Text = "" Or txtComments.Text = "" Then
    MsgBox "Cannot Save This Information, you have not filled in all fields, Please check and try again.", vbInformation, "The Genesis Project"
    Exit Sub
End If
    
    dtaDeathlink.Recordset.AddNew
    dtaDeathlink.Recordset.Update
    dtaDeathlink.Recordset.MoveLast
    txtForenamelink.Text = txtForename.Text
    txtSurenameLink.Text = txtSurename.Text
    txtAgeLink.Text = txtAge.Text
    txtCommentsLink.Text = txtComments.Text
    dtaDeathlink.Refresh
    MSFlexGrid1.Refresh
    
End Sub

Private Sub Form_Load()

    Me.Picture = frmPictureBase.imgBG.Picture
    Me.Width = 11235
    Me.Height = 5955

    apppath = App.Path
    dtaDeathlink.DatabaseName = apppath & "\KJV.mdb"
    dtaDeathlink.RecordSource = "DeathRecords"
     
End Sub
