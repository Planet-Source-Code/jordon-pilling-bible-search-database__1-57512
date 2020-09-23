VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H006B0500&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11310
   Icon            =   "frmHelp.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBacktomenu 
      Caption         =   "Back To Menu"
      Height          =   375
      Left            =   9720
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "- Help And Hints"
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
      Left            =   3960
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblHelp 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmHelp.frx":0442
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   2655
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   5895
   End
End
Attribute VB_Name = "frmhelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBacktomenu_Click()
    Unload Me
    ShapedForm.Show
End Sub

Private Sub Form_Load()
Me.Picture = frmPictureBase.imgBG.Picture
Me.Width = 11235
Me.Height = 5955
End Sub
