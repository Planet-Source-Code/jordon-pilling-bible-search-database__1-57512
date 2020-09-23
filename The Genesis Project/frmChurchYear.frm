VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmChurchYear 
   BackColor       =   &H006B0500&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5955
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11235
   LinkTopic       =   "Form1"
   ScaleHeight     =   5955
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraNavigation 
      BackColor       =   &H006B0500&
      Caption         =   "Navigation"
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
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   10935
      Begin VB.CommandButton cmdback 
         Caption         =   "&Back"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdForward 
         Caption         =   "&Forward"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdStop 
         Caption         =   "&Stop"
         Height          =   375
         Left            =   3720
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblPrompt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status: "
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
         Height          =   240
         Left            =   5040
         TabIndex        =   8
         Top             =   480
         Width           =   690
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loading ..."
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
         Height          =   240
         Left            =   5760
         TabIndex        =   7
         Top             =   480
         Width           =   915
      End
   End
   Begin VB.CommandButton cmdBacktomenu 
      Caption         =   "Back To Menu"
      Height          =   375
      Left            =   9600
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin SHDocVwCtl.WebBrowser webBrowser 
      Height          =   4215
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   10935
      ExtentX         =   19288
      ExtentY         =   7435
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "- The Church Year"
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
Attribute VB_Name = "frmChurchYear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBacktomenu_Click()
    Unload Me
    ShapedForm.Show
End Sub

Private Sub Form_Load()
    Me.Picture = frmPictureBase.imgBG.Picture
    Me.Width = 11235
    Me.Height = 5955
    MsgBox "To get the latest information, your computer must be connected to the internet, if you are not connected to the internet, this form will not display any information.", vbInformation, "The Genesis Project"
    webBrowser.Navigate "http://217.19.224.165/liturgyframe.htm"
End Sub

Private Sub cmdBack_Click()
    On Error Resume Next
    webBrowser.GoBack
End Sub

Private Sub cmdForward_Click()
    On Error Resume Next
    webBrowser.GoForward
End Sub

Private Sub cmdRefresh_Click()
    On Error Resume Next
    webBrowser.Refresh
End Sub

Private Sub cmdStop_Click()
    lblStatus.Caption = "Interrupted"
    webBrowser.Stop
End Sub

Private Sub webBrowser_DownloadBegin()
    lblStatus.Caption = "Loading ..."
End Sub

Private Sub webBrowser_DownloadComplete()
    lblStatus.Caption = "Loaded"
End Sub

