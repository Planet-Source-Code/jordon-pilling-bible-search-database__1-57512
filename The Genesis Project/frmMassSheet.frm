VERSION 5.00
Begin VB.Form frmMassSheet 
   BackColor       =   &H006B0500&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   11250
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtFinishedSheet 
      Height          =   4335
      Left            =   1080
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   9015
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "Confirm Changes"
      Height          =   615
      Left            =   9720
      TabIndex        =   39
      Top             =   3960
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txt1dump 
      Height          =   3285
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   38
      Top             =   1320
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.TextBox txt4dump 
      Height          =   3285
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   37
      Top             =   1320
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.TextBox txt2dummp 
      Height          =   3285
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   36
      Top             =   1320
      Visible         =   0   'False
      Width           =   7935
   End
   Begin VB.CommandButton cmdprint 
      Caption         =   "Print"
      Height          =   375
      Left            =   9600
      TabIndex        =   35
      Top             =   5400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtSelectedPrayer 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      DataField       =   "Prayer Text"
      DataSource      =   "databaselink"
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   33
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Data databaselink 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Avalon\My Documents\Higher National Diploma\College\Project\The Genesis Project\KJV.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Prayers"
      Top             =   240
      Visible         =   0   'False
      Width           =   2460
   End
   Begin VB.Frame fme 
      BackColor       =   &H006B0500&
      Caption         =   "Done!"
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
      Height          =   2055
      Index           =   4
      Left            =   2640
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdBuild 
         Caption         =   "Buid"
         Height          =   375
         Left            =   4440
         TabIndex        =   30
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackColor       =   &H006B0500&
         Caption         =   "OK, the wizard now has enough information to build your sheet. click Build, to display."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   5535
      End
   End
   Begin VB.Frame fme 
      BackColor       =   &H006B0500&
      Caption         =   "Step 4 of 4"
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
      Height          =   2055
      Index           =   3
      Left            =   2640
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   6135
      Begin VB.ComboBox cmbSimplelist 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1920
         TabIndex        =   32
         Top             =   1560
         Width           =   3015
      End
      Begin VB.CommandButton cmdConfirm 
         Caption         =   "Confirm"
         Height          =   375
         Left            =   5040
         TabIndex        =   29
         Top             =   1560
         Width           =   975
      End
      Begin VB.TextBox txtChurch 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2040
         TabIndex        =   28
         Top             =   960
         Width           =   3495
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         ForeColor       =   &H00E0E0E0&
         Height          =   285
         Left            =   2040
         TabIndex        =   27
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label lbl4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H006B0500&
         Caption         =   "Eucharistic Prayer:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   31
         Top             =   1560
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H006B0500&
         Caption         =   "Church Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   26
         Top             =   960
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H006B0500&
         Caption         =   "Youre Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
   End
   Begin VB.Frame fme 
      BackColor       =   &H006B0500&
      Caption         =   "Step 3 of 4"
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
      Height          =   2055
      Index           =   2
      Left            =   2640
      TabIndex        =   7
      Top             =   1920
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmdCompres 
         Caption         =   "Commpress Layout"
         Height          =   375
         Left            =   1920
         TabIndex        =   25
         Top             =   960
         Width           =   1455
      End
      Begin VB.CommandButton cmdFormally 
         Caption         =   "Formal Layout"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackColor       =   &H006B0500&
         Caption         =   "Compact Sheet, or lay out Formally?  Formal View will take up more paper."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   5655
      End
   End
   Begin VB.Frame fme 
      BackColor       =   &H006B0500&
      Caption         =   "Step 2 of 4 - Please Choose Youre Readings By Clicking Button"
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
      Height          =   2055
      Index           =   1
      Left            =   2640
      TabIndex        =   9
      Top             =   1920
      Visible         =   0   'False
      Width           =   6135
      Begin VB.CommandButton cmd3 
         Caption         =   "Tweak"
         Height          =   285
         Left            =   5160
         TabIndex        =   42
         Top             =   1560
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd2 
         Caption         =   "Tweak"
         Height          =   285
         Left            =   5160
         TabIndex        =   41
         Top             =   1200
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmd1 
         Caption         =   "Tweak"
         Height          =   285
         Left            =   5160
         TabIndex        =   40
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.CommandButton cmdGet3 
         Caption         =   ">"
         Height          =   285
         Left            =   4800
         TabIndex        =   22
         Top             =   1560
         Width           =   255
      End
      Begin VB.CommandButton cmdGet2 
         Caption         =   ">"
         Height          =   285
         Left            =   4800
         TabIndex        =   21
         Top             =   1200
         Width           =   255
      End
      Begin VB.CommandButton cmdGet1 
         Caption         =   ">"
         Height          =   285
         Left            =   4800
         TabIndex        =   20
         Top             =   840
         Width           =   255
      End
      Begin VB.TextBox txtGospel 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   1560
         Width           =   3495
      End
      Begin VB.TextBox txtSecond 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtFirst 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   840
         Width           =   3495
      End
      Begin VB.Timer tmrDateAndTime 
         Interval        =   1000
         Left            =   5880
         Top             =   1440
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   6000
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "1st Reading:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   -480
         TabIndex        =   16
         Top             =   840
         Width           =   1695
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "2nd Reading:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   -480
         TabIndex        =   15
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Gospel:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   -480
         TabIndex        =   14
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label lblTime 
         BackStyle       =   0  'Transparent
         Caption         =   "Time:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   2400
         TabIndex        =   13
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblDate 
         BackStyle       =   0  'Transparent
         Caption         =   "Date:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame fme 
      BackColor       =   &H006B0500&
      Caption         =   "Step 1 of 4"
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
      Height          =   2055
      Index           =   0
      Left            =   2640
      TabIndex        =   2
      Top             =   1920
      Width           =   6135
      Begin VB.Label Label1 
         BackColor       =   &H006B0500&
         Caption         =   $"frmMassSheet.frx":0000
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1455
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   5535
      End
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Start"
      Height          =   375
      Left            =   9600
      TabIndex        =   4
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Menu"
      Height          =   375
      Left            =   9600
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label lblPrayerName 
      DataField       =   "Prayer Name"
      DataSource      =   "databaselink"
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
      Left            =   240
      TabIndex        =   34
      Top             =   600
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "- Mass Sheet Wizard"
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
Attribute VB_Name = "frmMassSheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Currentstage As Integer
Dim newline As String
Dim completedpassage As String

Private Sub cmbSimplelist_Click()
chosenEP = cmbSimplelist
End Sub

Private Sub cmd1_Click()
    txt1dump.Visible = True
cmdEdit.Visible = True
End Sub

Private Sub cmd2_Click()
    txt2dummp.Visible = True
cmdEdit.Visible = True
End Sub

Private Sub cmd3_Click()
    txt4dump.Visible = True
    cmdEdit.Visible = True
End Sub

Private Sub cmdBack_Click()
    Unload Me
    ShapedForm.Show
End Sub

Private Sub cmdBuild_Click()
    newline = Chr(13) + Chr(10)
    cmdprint.Visible = True
    txtFinishedSheet.Visible = True
    
    completedpassage = "" & completedpassage & "" & newline & "Sheet Produced by:" & Priestname & "" & newline & ""
    completedpassage = "" & completedpassage & "" & newline & "For Church:" & Churchname & "" & newline & ""
searchstring = "Penitential Rite"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = "Mercy"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = "Gloria"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = "The profession of faith"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""

'Puts in 3 bible readings
    completedpassage = "" & completedpassage & "" & newline & "FIRST READING" & newline & "--------------------------" & newline & "A Reading from the book of:"
    completedpassage = "" & completedpassage & "" & newline & "" & txt1dump.Text & "PRIEST:This is the word of the lord" & newline & "RESPONSE:Thanks be to god" & newline & ""
    
    completedpassage = "" & completedpassage & "" & newline & "SECOND READING" & newline & "--------------------------" & newline & "A Reading from the book of:"
    completedpassage = "" & completedpassage & "" & newline & "" & txt2dummp.Text & "PRIEST:This is the word of the lord" & newline & "RESPONSE:Thanks be to god" & newline & ""
    
    completedpassage = "" & completedpassage & "" & newline & "GOSPEL" & newline & "--------------------------" & newline & "PRIEST:The Lord be with you" & newline & "RESPONSE:And also with you" & newline & "+A reading from the holy Gospel according to"
    completedpassage = "" & completedpassage & "" & newline & "" & txt4dump.Text & "This is the Gospel of the Lord" & newline & "Praise to you, Lord Jesus Christ." & newline & ""
        
searchstring = "Bidding Prayers"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = "Liturgy of the Eucharist"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = "Prayer over the gifts"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = "Euchristic Blessing"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = chosenEP
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = "Holy Holy"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = "Our Father (The Lords Prayer)"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = "Pre-Communion"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = "Lamb of God"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = "Consumation of the host"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = "Prayer after Communion"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""
searchstring = "Concluding Rite"
    searchprayers
    completedpassage = "" & completedpassage & "" & newline & "" & txtSelectedPrayer & "" & newline & ""

txtFinishedSheet.Text = "" & completedpassage & ""
End Sub
Sub searchprayers()
        databaselink.Recordset.MoveFirst
        searchwhat = "[Prayer Name] = '"
        finalsearch = searchwhat & searchstring & "'"
        databaselink.Recordset.FindNext (finalsearch)
End Sub
Private Sub cmdCompres_Click()
    layout = "compres"
    fme(2).Visible = False
    fme(3).Visible = True
    Currentstage = 4
End Sub

Private Sub cmdConfirm_Click()
If txtName.Text = "" Or txtChurch.Text = "" Or cmbSimplelist.Text = "" Then
    MsgBox "Please enter youre name and church name before you continue.", vbInformation, "The Genesis Project"
    Exit Sub
End If
    Priestname = txtName.Text
    Churchname = txtChurch.Text
    fme(3).Visible = False
    fme(4).Visible = True
End Sub

Private Sub cmdFormally_Click()
    layout = "formal"
    fme(2).Visible = False
    fme(3).Visible = True
    Currentstage = 4
End Sub

Private Sub cmdGet1_Click()
    callfor = 1
    frmBibleBase.Show
    cmd1.Visible = True
End Sub

Private Sub cmdGet2_Click()
    callfor = 2
    frmBibleBase.Show
    cmd2.Visible = True
End Sub

Private Sub cmdGet3_Click()
    callfor = 3
    frmBibleBase.Show
    cmd3.Visible = True
End Sub

Private Sub cmdNext_Click()

    txt1dump.Visible = False
    txt2dummp.Visible = False
    txt4dump.Visible = False
    cmdEdit.Visible = False

cmdNext.Caption = "Next >"

If Currentstage = 5 Then
    Currentstage = 0
    fme(0).Visible = True
    fme(1).Visible = False
    fme(2).Visible = False
    fme(3).Visible = False
    fme(4).Visible = False
End If
If Currentstage = 2 Then
    If txtFirst.Text = "" Or txtSecond.Text = "" Or txtGospel.Text = "" Then
        MsgBox "Cannot go to next stage yet, please choose three readings before you continue", vbInformation, "The Genesis Project"
        Exit Sub
    Else
        fme(Currentstage).Visible = True
        Currentstage = Currentstage + 1
        cmdNext.Visible = False
    End If
Else
    fme(Currentstage).Visible = True
    Currentstage = Currentstage + 1
End If


End Sub

Private Sub cmdprint_Click()
If txtFinishedSheet.Visible = False Then
    MsgBox "Please complete the wizard before printing.", vbInformation, "The Genesis Project"
    Exit Sub
End If
    Printer.Orientation = 2
    Printer.FontName = "Arial"
    Printer.FontSize = 8
    Printer.FontBold = False
    Printer.Print txtFinishedSheet.Text
    Printer.EndDoc
End Sub

Private Sub cmdEdit_Click()
    txt1dump.Visible = False
    txt2dummp.Visible = False
    txt4dump.Visible = False
    cmdEdit.Visible = False
End Sub

Private Sub Form_Load()
Me.Picture = frmPictureBase.imgBG.Picture
Me.Width = 11235
Me.Height = 5955
    apppath = App.Path
    databaselink.DatabaseName = apppath & "\KJV.mdb"
    databaselink.RecordSource = "Prayers"
    cmbSimplelist.AddItem "I"
    cmbSimplelist.AddItem "II"
    cmbSimplelist.AddItem "III"
    cmbSimplelist.AddItem "IV"
End Sub

Private Sub tmrDateAndTime_Timer()
    lblDate.Caption = "Date: " & Date & ""
    lblTime.Caption = "Time: " & Time & ""
End Sub
