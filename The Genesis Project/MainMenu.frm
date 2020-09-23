VERSION 5.00
Begin VB.Form frmBibleBase 
   BackColor       =   &H006B0500&
   BorderStyle     =   0  'None
   Caption         =   "Bible Base"
   ClientHeight    =   5970
   ClientLeft      =   0
   ClientTop       =   105
   ClientWidth     =   11250
   LinkTopic       =   "Form1"
   ScaleHeight     =   5970
   ScaleWidth      =   11250
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back To Menu"
      Height          =   375
      Left            =   9600
      TabIndex        =   30
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H006B0500&
      Caption         =   "Passage Viewer"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2640
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1695
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Book"
         DataSource      =   "Data1"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "BookTitle"
         DataSource      =   "Data1"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   2880
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Chapter"
         DataSource      =   "Data1"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         DataField       =   "Verse"
         DataSource      =   "Data1"
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Book No'"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   430
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Book Title:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   430
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Chapter:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   430
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Verse:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   5880
         TabIndex        =   7
         Top             =   430
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Passage:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   360
         TabIndex        =   6
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         BorderStyle     =   1  'Fixed Single
         DataField       =   "TextData"
         DataSource      =   "Data1"
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   1320
         TabIndex        =   5
         Top             =   720
         Width           =   5895
         WordWrap        =   -1  'True
      End
   End
   Begin VB.TextBox txtSelected 
      Appearance      =   0  'Flat
      BackColor       =   &H006B0500&
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   3840
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   26
      Top             =   4200
      Width           =   7215
   End
   Begin VB.ListBox lstProgress 
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
      ForeColor       =   &H00C0C0C0&
      Height          =   1500
      Left            =   120
      TabIndex        =   25
      Top             =   4200
      Width           =   3615
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H006B0500&
      Caption         =   "Bible Address"
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   120
      TabIndex        =   17
      Top             =   840
      Width           =   10935
      Begin VB.CommandButton cmdformasssheet 
         Height          =   390
         Left            =   4440
         TabIndex        =   32
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.TextBox txtInput 
         BackColor       =   &H006B0500&
         ForeColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   240
         TabIndex        =   29
         Text            =   "[Enter Bible Reference Here]"
         Top             =   720
         Width           =   4095
      End
      Begin VB.CommandButton cmdGetPassage 
         Caption         =   "Get Passage..."
         Height          =   390
         Left            =   4440
         TabIndex        =   28
         Top             =   720
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Start Over"
         Height          =   390
         Left            =   4440
         TabIndex        =   24
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CommandButton cmdGo 
         Caption         =   "Check For Errors"
         Height          =   390
         Left            =   4440
         TabIndex        =   18
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblHelp 
         BackStyle       =   0  'Transparent
         Caption         =   $"MainMenu.frx":0000
         ForeColor       =   &H00C0C0C0&
         Height          =   1455
         Left            =   7320
         TabIndex        =   27
         Top             =   200
         Width           =   3495
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Bible Reference Below:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H006B0500&
      Caption         =   "Find Passage"
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   10935
      Begin VB.ComboBox ListLineTo 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "MainMenu.frx":019D
         Left            =   6600
         List            =   "MainMenu.frx":019F
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   720
         Width           =   4095
      End
      Begin VB.ComboBox ListChapter 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   6600
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   360
         Width           =   4095
      End
      Begin VB.ComboBox ListLineFrom 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         Enabled         =   0   'False
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "MainMenu.frx":01A1
         Left            =   1320
         List            =   "MainMenu.frx":01A3
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   720
         Width           =   4215
      End
      Begin VB.ComboBox ListBooks 
         Appearance      =   0  'Flat
         BackColor       =   &H006B0500&
         ForeColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "MainMenu.frx":01A5
         Left            =   1320
         List            =   "MainMenu.frx":01A7
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "To Line:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "From Line:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Chapter:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5400
         TabIndex        =   14
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Book:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Data Data1 
      Appearance      =   0  'Flat
      Caption         =   "Data1"
      Connect         =   "Access 2000;"
      DatabaseName    =   "C:\Documents and Settings\Avalon Alpha\My Documents\HND Year 1\Project\The Genesis Project\KJV.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   270
      Left            =   120
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "BibleTable"
      Top             =   480
      Visible         =   0   'False
      Width           =   2130
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "- Bible Viewer"
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
      TabIndex        =   31
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label lblBuffer 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8160
      TabIndex        =   23
      Top             =   2640
      Width           =   3135
   End
End
Attribute VB_Name = "frmBibleBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim loopnumber As Integer
Dim AnErrorHasOccured As Boolean
Dim startline As Integer
Dim endline As Integer
Dim bookID, chapter, looptrigger, ErrorTrapping As Integer
Dim length As Integer
Dim tominus As Integer
Dim toBeSegmentedString, ReadInChar As String
Dim passage, newline As String
Dim apppath, booknumber, chapters, chapters2, StringSearch, SS2, SS3 As String

Dim Finalbook As String
Dim Finalchapter As String
Dim FinalFirstline As String
Dim FinalEndline As String

Private Sub cmdBack_Click()
    Unload Me
    ShapedForm.Show
End Sub

Private Sub cmdformasssheet_Click()
searchBible
Select Case callfor

    Case 1
        frmMassSheet.txtFirst.Text = txtInput.Text
        frmMassSheet.txt1dump.Text = txtSelected.Text
    Case 2
        frmMassSheet.txtSecond.Text = txtInput.Text
        frmMassSheet.txt2dummp.Text = txtSelected.Text
    Case 3
        frmMassSheet.txtGospel.Text = txtInput.Text
        frmMassSheet.txt4dump.Text = txtSelected.Text
End Select

Unload Me

End Sub

Private Sub cmdGetPassage_Click()

    searchBible


End Sub
Sub searchBible()

    newline = Chr(13) + Chr(10)
    passage = "" & SS3 & "" & newline & ""
    
Data1.Recordset.MoveFirst
Data1.RecordSource = "SELECT * FROM BibleTable WHERE BookTitle = '" & Finalbook & "' AND Chapter =" & Finalchapter
Data1.Refresh
Data1.Recordset.MoveLast
loopnumber = FinalFirstline

Do
   'On Error GoTo 20:
        Data1.Recordset.MoveFirst
        searchstring = loopnumber
        searchwhat = "[Verse] = '"
        finalsearch = searchwhat & searchstring & "'"
        Data1.Recordset.FindNext (finalsearch)
        passage = "" & passage & "" & Label6.Caption & "" & newline & ""
        loopnumber = loopnumber + 1

Loop Until loopnumber > FinalEndline
    
txtSelected.Text = "" & passage & ""
    
20:
End Sub
Private Sub cmdGo_Click()

If txtInput.Text = "" Or txtInput.Text = "[Enter Bible Reference Here]" Then
    MsgBox "Please either enter a bible reference or use the list boxes to help you choose youre passage.", vbInformation, "The Genesis Project"
    Exit Sub
End If
'This part of code will split an entered query and find relevant data and display it

'An entered string such as "Mark 23(25:34) could be entered, the following code will read the mark into a variable
'Get Book name
tominus = 0
StringSearch = txtInput.Text
SS3 = StringSearch
ErrorTrapping = 0
AnErrorHasOccured = False
length = Len(StringSearch)
lstProgress.AddItem "Proccessing String..."
lstProgress.AddItem "Getting Book..."
Finalbook = ""
looptrigger = 1
Do
    ReadInChar = Mid(StringSearch, looptrigger, 1)
    looptrigger = looptrigger + 1
    ErrorTrapping = ErrorTrapping + 1
    If ReadInChar <> " " Then
        Finalbook = "" & Finalbook & "" & ReadInChar & ""
    Else
        tominus = Len(Finalbook)

        StringSearch = Right(StringSearch, length - tominus - 1)
        StringSearch = Trim(StringSearch)
        lstProgress.AddItem "Book determined as: " & Finalbook & ""
    End If
    If ErrorTrapping >= 20 Then GoTo Read_Error:
Loop Until ReadInChar = " "
    checkspelling
    If AnErrorHasOccured = True Then GoTo Spelling_Error:
    
'Get Chapter Number
length = Len(StringSearch)
lstProgress.AddItem "Getting Chapter..."
ErrorTrapping = 0
looptrigger = 1
Finalchapter = ""
Do
    ReadInChar = Mid(StringSearch, looptrigger, 1)
    looptrigger = looptrigger + 1
    ErrorTrapping = ErrorTrapping + 1
    If ReadInChar <> "(" Then
        Finalchapter = "" & Finalchapter & "" & ReadInChar & ""
    Else
        tominus = Len(Finalchapter)

        StringSearch = Right(StringSearch, length - tominus - 1)
        StringSearch = Trim(StringSearch)
        lstProgress.AddItem "Chapter determined as: " & Finalchapter & ""
    End If
    If ErrorTrapping >= 20 Then GoTo Read_Error:
Loop Until ReadInChar = "("

'Get Start Line
length = Len(StringSearch)
lstProgress.AddItem "Getting Line ref for start..."
ErrorTrapping = 0
looptrigger = 1
FinalFirstline = ""
Do
    ReadInChar = Mid(StringSearch, looptrigger, 1)
    looptrigger = looptrigger + 1
    ErrorTrapping = ErrorTrapping + 1
    If ReadInChar <> ":" Then
        FinalFirstline = "" & FinalFirstline & "" & ReadInChar & ""
    Else
        tominus = Len(FinalFirstline)

        StringSearch = Right(StringSearch, length - tominus - 1)
        StringSearch = Trim(StringSearch)
        lstProgress.AddItem "Start Line determined as: " & FinalFirstline & ""
    End If
    If ErrorTrapping >= 20 Then GoTo Read_Error:
Loop Until ReadInChar = ":"

'Get end line
length = Len(StringSearch)
lstProgress.AddItem "Getting Line ref for start..."
ErrorTrapping = 0
looptrigger = 1
FinalEndline = ""
Do
    ReadInChar = Mid(StringSearch, looptrigger, 1)
    looptrigger = looptrigger + 1
    ErrorTrapping = ErrorTrapping + 1
    If ReadInChar <> ")" Then
        FinalEndline = "" & FinalEndline & "" & ReadInChar & ""
    Else
        tominus = Len(FinalEndline)

        StringSearch = Right(StringSearch, length - tominus - 1)
        StringSearch = Trim(StringSearch)
        lstProgress.AddItem "End Line determined as: " & FinalEndline & ""
    End If
    If ErrorTrapping >= 20 Then GoTo Read_Error:
Loop Until ReadInChar = ")"

If Int(FinalEndline) < Int(FinalFirstline) Then GoTo cannot_read_backwards:
    If callfor = 0 Then
        cmdGetPassage.Visible = True
    Else
        cmdformasssheet.Visible = True
    End If
tominus = 0
Exit Sub

Read_Error:
        clearform
        lstProgress.AddItem "ERROR: Check Search Format"
        clean_Variables
Exit Sub

Spelling_Error:
        clearform
        lstProgress.AddItem "ERROR: Book does not exist in bible"
        clean_Variables
Exit Sub

cannot_read_backwards:
        clearform
        lstProgress.AddItem "ERROR: Cannot Read Backwards, check lines"
        clean_Variables
End Sub

Sub clean_Variables()

    Finalbook = ""
    Finalchapter = 0
    FinalFirstline = 0
    FinalEndline = 0
    StringSearch = ""
    SS3 = ""
    SS2 = ""
    ErrorTrapping = 0

End Sub
Private Sub cmdClear_Click()
    clearform
End Sub

Sub clearform()
    ListLineTo.Clear
    ListLineTo.Enabled = False
    ListLineFrom.Clear
    ListLineFrom.Enabled = False
    ListChapter.Clear
    ListChapter.Enabled = False
    ListBooks.Clear
    ListBooks.Enabled = True
    txtInput.Text = ""
    StringSearch = ""
    fillinbooks
    clean_Variables
    lstProgress.Clear
End Sub
Sub fillinbooks()

    ListBooks.AddItem "Genesis"
    ListBooks.AddItem "Exodus"
    ListBooks.AddItem "Leviticus"
    ListBooks.AddItem "Numbers"
    ListBooks.AddItem "Deuteronomy"
    ListBooks.AddItem "Joshua"
    ListBooks.AddItem "Judges"
    ListBooks.AddItem "Ruth"
    ListBooks.AddItem "1 Samuel"
    ListBooks.AddItem "2 Samuel"
    ListBooks.AddItem "1 Kings"
    ListBooks.AddItem "2 Kings"
    ListBooks.AddItem "1 Chronicles"
    ListBooks.AddItem "2 Chronicles"
    ListBooks.AddItem "Ezra"
    ListBooks.AddItem "Nehemiah"
    ListBooks.AddItem "Esther"
    ListBooks.AddItem "Job"
    ListBooks.AddItem "Psalms"
    ListBooks.AddItem "Proverbs"
    ListBooks.AddItem "Ecclesiastes"
    ListBooks.AddItem "Song of Solomon"
    ListBooks.AddItem "Isaiah"
    ListBooks.AddItem "Jeremiah"
    ListBooks.AddItem "Lamentations"
    ListBooks.AddItem "Ezekiel"
    ListBooks.AddItem "Daniel"
    ListBooks.AddItem "Hosea"
    ListBooks.AddItem "Joel"
    ListBooks.AddItem "Amos"
    ListBooks.AddItem "Obadiah"
    ListBooks.AddItem "Jonah"
    ListBooks.AddItem "Micah"
    ListBooks.AddItem "Nahum"
    ListBooks.AddItem "Habakkuk"
    ListBooks.AddItem "Zephaniah"
    ListBooks.AddItem "Haggai"
    ListBooks.AddItem "Zechariah"
    ListBooks.AddItem "Malachi"
    ListBooks.AddItem "Matthew"
    ListBooks.AddItem "Mark"
    ListBooks.AddItem "Luke"
    ListBooks.AddItem "John"
    ListBooks.AddItem "Acts"
    ListBooks.AddItem "Romans"
    ListBooks.AddItem "1 Corinthians"
    ListBooks.AddItem "2 Corinthians"
    ListBooks.AddItem "Galatians"
    ListBooks.AddItem "Ephesians"
    ListBooks.AddItem "Philippians"
    ListBooks.AddItem "Colossians"
    ListBooks.AddItem "1 Thessalonians"
    ListBooks.AddItem "2 Thessalonians"
    ListBooks.AddItem "1 Timothy"
    ListBooks.AddItem "2 Timothy"
    ListBooks.AddItem "Titus"
    ListBooks.AddItem "Philemon"
    ListBooks.AddItem "Hebrews"
    ListBooks.AddItem "James"
    ListBooks.AddItem "1 Peter"
    ListBooks.AddItem "2 Peter"
    ListBooks.AddItem "1 John"
    ListBooks.AddItem "2 John"
    ListBooks.AddItem "3 John"
    ListBooks.AddItem "Jude"
    ListBooks.AddItem "Revelation"
End Sub

Private Sub Form_Load()
    Me.Picture = frmPictureBase.imgBG.Picture
    fillinbooks
    apppath = App.Path
    Data1.DatabaseName = apppath & "\KJV.mdb"
    If callfor <> 0 Then
        cmdformasssheet.Caption = "Use for reading " & callfor & ""
    End If
End Sub

Private Sub ListBooks_Click()

'I would use SQL to do this part if i had enough time to sit and tweak it, but i dont,
'...so i have done it this way because i know it will work even though it is bulky code

    StringSearch = ListBooks
    txtInput.Text = "" & StringSearch & ""
    booknumber = ListBooks.ListIndex + 1
    bookID = ListBooks.ListIndex + 1
    
    Select Case ListBooks.ListIndex

    Case 0              'Genesis
        chapters = 50
    Case 1              'Exodus
        chapters = 40
    Case 2              'Leviticus
        chapters = 27
    Case 3              'Numbers
        chapters = 36
    Case 4              'Deuteronomy
        chapters = 34
    Case 5              'Joshua
        chapters = 24
    Case 6              'Judges
        chapters = 21
    Case 7              'Ruth
        chapters = 4
    Case 8              '1 Samuel
        chapters = 31
    Case 9              '2 Samuel
        chapters = 24
    Case 10             '1 Kings
        chapters = 22
    Case 11             '2 Kings
        chapters = 25
    Case 12             '1 Chronicles
        chapters = 29
    Case 13             '2 Chronicles
        chapters = 36
    Case 14             'Ezra
        chapters = 10
    Case 15             'Nehemiah
        chapters = 13
    Case 16             'Esther
        chapters = 10
    Case 17             'Job
        chapters = 42
    Case 18             'Psalms
        chapters = 150
    Case 19             'Proverbs
        chapters = 31
    Case 20             'Ecclesiastes
        chapters = 12
    Case 21             'Song of Solomon
        chapters = 8
    Case 22             'Isaiah
        chapters = 66
    Case 23             'Jeremiah
        chapters = 52
    Case 24             'Lamentations
        chapters = 5
    Case 25             'Ezekiel
        chapters = 48
    Case 26             'Daniel
        chapters = 12
    Case 27             'Hosea
        chapters = 14
    Case 28             'Joel
        chapters = 3
    Case 29             'Amos
        chapters = 9
    Case 30             'Obadiah
        chapters = 1
    Case 31             'Jonah
        chapters = 4
    Case 32             'Micah
        chapters = 7
    Case 33             'Nahum
        chapters = 3
    Case 34             'Habakkuk
        chapters = 3
    Case 35             'Zephaniah
        chapters = 3
    Case 36             'Haggai
        chapters = 2
    Case 37             'Zechariah
        chapters = 14
    Case 38             'Malachi
        chapters = 4
    Case 39             'Matthew
        chapters = 28
    Case 40             'Mark
        chapters = 16
    Case 41             'Luke
        chapters = 24
    Case 42             'John
        chapters = 21
    Case 43             'Acts
        chapters = 28
    Case 44             'Romans
        chapters = 16
    Case 45             '1 Corinthians
        chapters = 16
    Case 46             '2 Corinthians
        chapters = 13
    Case 47             'Galatians
        chapters = 6
    Case 48             'Ephesians
        chapters = 6
    Case 49             'Philippians
        chapters = 4
    Case 50             'Colossians
        chapters = 4
    Case 51             '1 Thessalonians
        chapters = 5
    Case 52             '2 Thessalonians
        chapters = 3
    Case 53             '1 Timothy
        chapters = 6
    Case 54             '2 Timothy
        chapters = 4
    Case 55             'Titus
        chapters = 3
    Case 56             'Philemon
        chapters = 1
    Case 57             'Hebrews
        chapters = 13
    Case 58             'James
        chapters = 5
    Case 59             '1 Peter
        chapters = 5
    Case 60             '2 Peter
        chapters = 3
    Case 61             '1 John
        chapters = 5
    Case 62             '2 John
        chapters = 1
    Case 63             '3 John
        chapters = 1
    Case 64             'Jude
        chapters = 1
    Case 65             'Revelation
        chapters = 22
End Select
chapters2 = chapters
fillchapters
ListChapter.Enabled = True
ListBooks.Enabled = False
End Sub

Sub fillchapters()
ListChapter.Clear
Do
    ListChapter.AddItem "" & chapters & ""
    chapters = chapters - 1
Loop Until chapters = 0


End Sub

Private Sub ListChapter_Click()

StringSearch = "" & StringSearch & " " & ListChapter & "("
chapter = ListChapter
txtInput.Text = StringSearch

Data1.Recordset.MoveFirst
Data1.RecordSource = "SELECT * FROM BibleTable WHERE BookTitle = '" & ListBooks & "' AND Chapter =" & ListChapter
Data1.Refresh
Data1.Recordset.MoveLast

chapters = Data1.Recordset.RecordCount
fill_lines
ListChapter.Enabled = False
ListLineFrom.Enabled = True

Data1.RecordSource = "SELECT * FROM BibleTable"
Data1.Refresh

End Sub
Sub fill_lines()

ListLineFrom.Clear
ListLineTo.Clear

Do
    ListLineFrom.AddItem "" & chapters & ""
    ListLineTo.AddItem "" & chapters & ""
    chapters = chapters - 1
Loop Until chapters = 0


End Sub

Private Sub ListLineFrom_Click()
    SS2 = StringSearch
    StringSearch = "" & StringSearch & "" & ListLineFrom & ":"
    startline = ListLineFrom
    txtInput.Text = StringSearch
    ListLineFrom.Enabled = False
    ListLineTo.Enabled = True
End Sub

Private Sub ListLineTo_Click()
endline = ListLineTo

    If startline > endline Then
        MsgBox "You have chosen to read backwards, please check youre lines!", vbInformation, "The Genesis Project"
        StringSearch = SS2
        txtInput.Text = StringSearch
        ListLineTo.Enabled = False
        ListLineFrom.Enabled = True
    Else
        StringSearch = "" & StringSearch & "" & ListLineTo & ")"
        txtInput.Text = StringSearch
        ListLineTo.Enabled = False
    End If
    
End Sub

Sub checkspelling()
Select Case Finalbook
    Case "Genesis"
    Case "Exodus"
    Case "Leviticus"
    Case "Numbers"
    Case "Deuteronomy"
    Case "Joshua"
    Case "Judges"
    Case "Ruth"
    Case "1 Samuel"
    Case "2 Samuel"
    Case "1 Kings"
    Case "2 Kings"
    Case "1 Chronicles"
    Case "2 Chronicles"
    Case "Ezra"
    Case "Nehemiah"
    Case "Esther"
    Case "Job"
    Case "Psalms"
    Case "Proverbs"
    Case "Ecclesiastes"
    Case "Song of Solomon"
    Case "Isaiah"
    Case "Jeremiah"
    Case "Lamentations"
    Case "Ezekiel"
    Case "Daniel"
    Case "Hosea"
    Case "Joel"
    Case "Amos"
    Case "Obadiah"
    Case "Jonah"
    Case "Micah"
    Case "Nahum"
    Case "Habakkuk"
    Case "Zephaniah"
    Case "Haggai"
    Case "Zechariah"
    Case "Malachi"
    Case "Matthew"
    Case "Mark"
    Case "Luke"
    Case "John"
    Case "Acts"
    Case "Romans"
    Case "1 Corinthians"
    Case "2 Corinthians"
    Case "Galatians"
    Case "Ephesians"
    Case "Philippians"
    Case "Colossians"
    Case "1 Thessalonians"
    Case "2 Thessalonians"
    Case "1 Timothy"
    Case "2 Timothy"
    Case "Titus"
    Case "Philemon"
    Case "Hebrews"
    Case "James"
    Case "1 Peter"
    Case "2 Peter"
    Case "1 John"
    Case "2 John"
    Case "3 John"
    Case "Jude"
    Case "Revelation"
    Case Else
        AnErrorHasOccured = True
End Select
End Sub

Private Sub txtinput_GotFocus()
    txtInput.Text = ""
End Sub
