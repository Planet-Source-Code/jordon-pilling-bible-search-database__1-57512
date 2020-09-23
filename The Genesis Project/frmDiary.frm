VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmDiary 
   BackColor       =   &H006B0500&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7575
   ClientLeft      =   1890
   ClientTop       =   1140
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   7575
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmSearch 
      BackColor       =   &H006B0500&
      Caption         =   "Search Results"
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
      Left            =   2280
      TabIndex        =   18
      Top             =   960
      Visible         =   0   'False
      Width           =   8775
      Begin VB.CommandButton cmdEndSearch 
         Caption         =   "Finish Search"
         Height          =   375
         Left            =   120
         TabIndex        =   27
         Top             =   1680
         Width           =   1095
      End
      Begin VB.CheckBox txtyesnolink 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Delete After Use"
         DataSource      =   "dtaDiary"
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2880
         TabIndex        =   26
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtmessagelink 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Message"
         DataSource      =   "dtaDiary"
         ForeColor       =   &H00E0E0E0&
         Height          =   645
         Left            =   2880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   25
         Top             =   1440
         Width           =   5775
      End
      Begin VB.TextBox txtsubjectlink 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Subject"
         DataSource      =   "dtaDiary"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2880
         TabIndex        =   24
         Top             =   1080
         Width           =   5775
      End
      Begin VB.TextBox txtremindlink 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Reminder On Date"
         DataSource      =   "dtaDiary"
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2880
         TabIndex        =   23
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remind Me On Date:"
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
         Left            =   840
         TabIndex        =   22
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Message:"
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
         Left            =   1680
         TabIndex        =   21
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
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
         Left            =   1920
         TabIndex        =   20
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Entry After Reminder:"
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
         Left            =   240
         TabIndex        =   19
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Timer tmrDateAndTime 
      Interval        =   1000
      Left            =   1200
      Top             =   720
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H006B0500&
      Caption         =   "Todays Information"
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
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   2055
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame2 
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
      Height          =   2415
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   10935
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Bindings        =   "frmDiary.frx":0000
         Height          =   1815
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   3201
         _Version        =   393216
         BackColorBkg    =   7013632
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   ""
      End
   End
   Begin VB.Data dtaDiary 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Avalon\My Documents\Higher National Diploma\College\Project\The Genesis Project\KJV.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   270
      Left            =   4800
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Diary"
      Top             =   120
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "Find By Subject"
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   2760
      Width           =   2055
   End
   Begin VB.CommandButton cmdSaveNew 
      Caption         =   "Save New Entry"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   2280
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H006B0500&
      Caption         =   "New Diary Entry"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   960
      Width           =   8775
      Begin VB.CheckBox chktruefalse 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2880
         TabIndex        =   10
         Top             =   720
         Width           =   375
      End
      Begin VB.TextBox txtMessage 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         ForeColor       =   &H00E0E0E0&
         Height          =   645
         Left            =   2880
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1440
         Width           =   5775
      End
      Begin VB.TextBox txtSubject 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2880
         TabIndex        =   8
         Top             =   1080
         Width           =   5775
      End
      Begin VB.TextBox txtRemindDate 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   5775
      End
      Begin VB.Label lblHeader 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Delete Entry After Reminder:"
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
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   2535
      End
      Begin VB.Label lblSubjest 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
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
         Left            =   1920
         TabIndex        =   6
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label lblMessage 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Message:"
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
         Left            =   1680
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label lblRemindMe 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Remind Me On Date:"
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
         Left            =   840
         TabIndex        =   4
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Menu"
      Height          =   375
      Left            =   9600
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "- Diary"
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
      Width           =   4335
   End
End
Attribute VB_Name = "frmDiary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
    Unload Me
    ShapedForm.Show
End Sub

Private Sub cmdEndSearch_Click()
    frmSearch.Visible = False
End Sub

Private Sub cmdFind_Click()
        frmInputbox.Show vbModal
        frmSearch.Visible = True
        dtaDiary.Recordset.MoveFirst
        searchstring = inputboxreturn
        searchwhat = "[Subject] = '"
        finalsearch = searchwhat & searchstring & "'"
        dtaDiary.Recordset.FindNext (finalsearch)
End Sub

Private Sub cmdSaveNew_Click()
    If txtRemindDate.Text = "" Or txtSubject.Text = "" Or txtMessage.Text = "" Then
    MsgBox "Cannot Save This Information, you have not filled in all fields, Please check and try again.", vbInformation, "The Genesis Project"
    Exit Sub
End If
    
    dtaDiary.Recordset.AddNew
    dtaDiary.Recordset.Update
    dtaDiary.Recordset.MoveLast
    txtremindlink.Text = txtRemindDate.Text
    txtyesnolink.Value = chktruefalse.Value
    txtsubjectlink.Text = txtSubject.Text
    txtmessagelink.Text = txtMessage.Text
    dtaDiary.Refresh
    MSFlexGrid1.Refresh
    txtRemindDate.Text = ""
    chktruefalse.Value = False
    txtSubject.Text = ""
    txtMessage.Text = ""
End Sub

Private Sub Form_Load()
    Me.Picture = frmPictureBase.imgBG.Picture
    Me.Width = 11235
    Me.Height = 5955
    
    apppath = App.Path
    dtaDiary.DatabaseName = apppath & "\KJV.mdb"
    dtaDiary.RecordSource = "Diary"
End Sub

Private Sub tmrDateAndTime_Timer()
    lblDate.Caption = "Date: " & Date & ""
    lblTime.Caption = "Time: " & Time & ""
End Sub
