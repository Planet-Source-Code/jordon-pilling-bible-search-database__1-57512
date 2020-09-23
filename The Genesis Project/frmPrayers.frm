VERSION 5.00
Begin VB.Form frmPrayers 
   BackColor       =   &H006B0500&
   BorderStyle     =   0  'None
   Caption         =   "Common Prayer"
   ClientHeight    =   5865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FmeAddNew 
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
      Height          =   3255
      Left            =   240
      TabIndex        =   16
      Top             =   2400
      Visible         =   0   'False
      Width           =   10935
      Begin VB.TextBox txtPrayerName 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   1800
         TabIndex        =   19
         Top             =   480
         Width           =   8895
      End
      Begin VB.TextBox txtPrayer 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         ForeColor       =   &H00E0E0E0&
         Height          =   2205
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   840
         Width           =   8895
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   255
         Left            =   480
         TabIndex        =   17
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Prayer Name:"
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
         TabIndex        =   21
         Top             =   480
         Width           =   1215
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Prayer Text:"
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
         TabIndex        =   20
         Top             =   840
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H006B0500&
      Caption         =   "Options"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   6600
      TabIndex        =   9
      Top             =   840
      Width           =   4575
      Begin VB.CommandButton cmdRefresh 
         Caption         =   ">"
         Height          =   255
         Left            =   3480
         TabIndex        =   12
         Top             =   1080
         Width           =   495
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   ">"
         Height          =   255
         Left            =   3480
         TabIndex        =   11
         Top             =   720
         Width           =   495
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   ">"
         Height          =   255
         Left            =   3480
         TabIndex        =   10
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Add a new prayer to database..........."
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
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   3375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Refresh Database............................."
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
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1080
         Width           =   3375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Current Prayer From Database:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.TextBox txtSearchQuery 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1080
      TabIndex        =   7
      Top             =   1320
      Width           =   4215
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "Search"
      Height          =   285
      Left            =   5400
      TabIndex        =   6
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ComboBox ListPrayers 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      ItemData        =   "frmPrayers.frx":0000
      Left            =   1080
      List            =   "frmPrayers.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   1800
      Width           =   4215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H006B0500&
      Caption         =   "Prayer:"
      ForeColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   240
      TabIndex        =   2
      Top             =   2400
      Width           =   10935
      Begin VB.TextBox txtSelectedPrayer 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Prayer Text"
         DataSource      =   "dtaPrayers"
         ForeColor       =   &H00FFFFFF&
         Height          =   2415
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         Top             =   720
         Width           =   10695
      End
      Begin VB.Label lblPrayerName 
         BackStyle       =   0  'Transparent
         DataField       =   "Prayer Name"
         DataSource      =   "dtaPrayers"
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
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   3375
      End
   End
   Begin VB.Data dtaPrayers 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Avalon\My Documents\Higher National Diploma\College\Project\The Genesis Project\KJV.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Prayers"
      Top             =   5520
      Visible         =   0   'False
      Width           =   2940
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Menu"
      Height          =   375
      Left            =   9600
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Search:"
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
      Height          =   255
      Left            =   360
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Prayer:"
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
      Height          =   255
      Left            =   -120
      TabIndex        =   5
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "- Common Prayer"
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
      TabIndex        =   1
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmPrayers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    Unload Me
    ShapedForm.Show
End Sub

Private Sub cmdButton_Click()
        dtaPrayers.Recordset.MoveFirst
        searchstring = txtSearchQuery.Text
        searchwhat = "[Prayer Name] = '"
        finalsearch = searchwhat & searchstring & "'"
        dtaPrayers.Recordset.FindNext (finalsearch)
End Sub

Private Sub cmdDelete_Click()
reply = MsgBox("Are you Sure you want to delete the following Prayer:" & vbCrLf & "" & vbCrLf & "" & txtSelectedPrayer.Text & "", vbYesNo, "The Genesis Project")
    If reply = 6 Then
        dtaPrayers.Recordset.Delete
        dtaPrayers.Recordset.MoveFirst
    End If

End Sub

Private Sub cmdRefresh_Click()
On Error GoTo 20:
    dtaPrayers.Recordset.Update
    dtaPrayers.Refresh
20:
End Sub

Private Sub cmdSave_Click()
    If txtPrayerName.Text = "" Or txtPrayer.Text = "" Then
    MsgBox "Cannot Save This Information, you have not filled in all fields, Please check and try again.", vbInformation, "The Genesis Project"
    Exit Sub
End If
    
    dtaPrayers.Recordset.AddNew
    dtaPrayers.Recordset.Update
    dtaPrayers.Recordset.MoveLast
    txtSelectedPrayer.Text = txtPrayer.Text
    lblPrayerName.Caption = txtPrayerName.Text
    dtaPrayers.Refresh
    txtSelectedPrayer.Text = ""
    lblPrayerName.Caption = ""
    FmeAddNew.Visible = False
End Sub

Private Sub cmdAdd_Click()
    FmeAddNew.Visible = True
End Sub

Private Sub Form_Load()
    Me.Picture = frmPictureBase.imgBG.Picture
    Me.Width = 11235
    Me.Height = 5955
    
    apppath = App.Path
    dtaPrayers.DatabaseName = apppath & "\KJV.mdb"
    dtaPrayers.RecordSource = "Prayers"
    
    ListPrayers.AddItem "Euchristic Prayer 1"
    ListPrayers.AddItem "Euchristic Prayer 2"
    ListPrayers.AddItem "Euchristic Prayer 3"
    ListPrayers.AddItem "Euchristic Prayer 4"
    ListPrayers.AddItem "Penitential Rite"
    ListPrayers.AddItem "Mercy"
    ListPrayers.AddItem "Gloria"
    ListPrayers.AddItem "The profession of faith"
    ListPrayers.AddItem "Bidding Prayers"
    ListPrayers.AddItem "Liturgy of the Eucharist"
    ListPrayers.AddItem "Prayer over the gifts"
    ListPrayers.AddItem "Euchristic Blessing"
    ListPrayers.AddItem "Holy Holy"
    ListPrayers.AddItem "Our Father (The Lords Prayer)"
    ListPrayers.AddItem "Pre-Communion"
    ListPrayers.AddItem "Lamb of God"
    ListPrayers.AddItem "Consumation of the host"
    ListPrayers.AddItem "Prayer after Communion"
    ListPrayers.AddItem "Concluding Rite"

End Sub

Private Sub ListPrayers_Click()

        dtaPrayers.Recordset.MoveFirst
        searchstring = ListPrayers
        searchwhat = "[Prayer Name] = '"
        finalsearch = searchwhat & searchstring & "'"
        dtaPrayers.Recordset.FindNext (finalsearch)
End Sub
