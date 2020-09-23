VERSION 5.00
Begin VB.Form ShapedForm 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "ShapedForm"
   ClientHeight    =   9090
   ClientLeft      =   3225
   ClientTop       =   1200
   ClientWidth     =   8910
   ControlBox      =   0   'False
   MaxButton       =   0   'False
   Picture         =   "Cross.frx":0000
   ScaleHeight     =   606
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   594
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Quit"
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
      Height          =   375
      Index           =   7
      Left            =   7440
      TabIndex        =   10
      Top             =   4320
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Bible"
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
      Height          =   375
      Index           =   6
      Left            =   7440
      TabIndex        =   9
      Top             =   3600
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Deaths"
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
      Height          =   375
      Index           =   5
      Left            =   7440
      TabIndex        =   8
      Top             =   3000
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Help"
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
      Height          =   375
      Index           =   4
      Left            =   7440
      TabIndex        =   7
      Top             =   2280
      Width           =   495
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Sheet"
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
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   6
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Internet"
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
      Height          =   375
      Index           =   2
      Left            =   960
      TabIndex        =   5
      Top             =   3720
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Prayers"
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
      Height          =   375
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Diary"
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
      Height          =   375
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "By J.Pilling @ The Phoenix Studios"
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
      Height          =   375
      Left            =   3120
      TabIndex        =   2
      Top             =   8520
      Width           =   2655
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblChosenOption 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Mass Sheet"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2895
      Left            =   3360
      TabIndex        =   1
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Image Button 
      Height          =   420
      Index           =   9
      Left            =   480
      Picture         =   "Cross.frx":107F54
      Top             =   0
      Width           =   405
   End
   Begin VB.Image Button 
      Height          =   420
      Index           =   8
      Left            =   0
      Picture         =   "Cross.frx":10849E
      Top             =   0
      Width           =   405
   End
   Begin VB.Image Button 
      Height          =   420
      Index           =   7
      Left            =   8040
      Picture         =   "Cross.frx":108A07
      Top             =   2040
      Width           =   405
   End
   Begin VB.Image Button 
      Height          =   420
      Index           =   6
      Left            =   8040
      Picture         =   "Cross.frx":108F51
      Top             =   2880
      Width           =   405
   End
   Begin VB.Image Button 
      Height          =   420
      Index           =   5
      Left            =   8040
      Picture         =   "Cross.frx":10949B
      Top             =   3600
      Width           =   405
   End
   Begin VB.Image Button 
      Height          =   420
      Index           =   4
      Left            =   8040
      Picture         =   "Cross.frx":1099E5
      Top             =   4440
      Width           =   405
   End
   Begin VB.Image Button 
      Height          =   420
      Index           =   3
      Left            =   450
      Picture         =   "Cross.frx":109F2F
      Top             =   4440
      Width           =   405
   End
   Begin VB.Image Button 
      Height          =   420
      Index           =   2
      Left            =   450
      Picture         =   "Cross.frx":10A479
      Top             =   3600
      Width           =   405
   End
   Begin VB.Image Button 
      Height          =   420
      Index           =   1
      Left            =   450
      Picture         =   "Cross.frx":10A9C3
      Top             =   2880
      Width           =   405
   End
   Begin VB.Image Button 
      Height          =   420
      Index           =   0
      Left            =   450
      Picture         =   "Cross.frx":10AF0D
      Top             =   2040
      Width           =   405
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Genesis Project"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "ShapedForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Type POINTAPI
   X As Long
   Y As Long
End Type
Private Const RGN_COPY = 5
Private Const CreatedBy = "VBSFC 7.0"
Private Const RegisteredTo = "Not Registered"
Private ResultRegion As Long
Private Function CreateFormRegion(ScaleX As Single, ScaleY As Single, OffsetX As Integer, OffsetY As Integer) As Long
    Dim HolderRegion As Long, ObjectRegion As Long, nRet As Long, counter As Integer
    Dim PolyPoints() As POINTAPI
    Dim STPPX As Integer, STPPY As Integer
    STPPX = Screen.TwipsPerPixelX
    STPPY = Screen.TwipsPerPixelY
    ResultRegion = CreateRectRgn(0, 0, 0, 0)
    HolderRegion = CreateRectRgn(0, 0, 0, 0)
    
    ReDim PolyPoints(0 To 309)
    For counter = 0 To 309
        PolyPoints(counter).X = GP0X(counter) * ScaleX * 15 / STPPX + OffsetX
        PolyPoints(counter).Y = GP0Y(counter) * ScaleY * 15 / STPPY + OffsetY
    Next counter
    ObjectRegion = CreatePolygonRgn(PolyPoints(0), 310, 1)
    nRet = CombineRgn(ResultRegion, ObjectRegion, ObjectRegion, RGN_COPY)
    DeleteObject ObjectRegion
    DeleteObject HolderRegion
    CreateFormRegion = ResultRegion
End Function
Private Function GP0X(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0X = 203
    Case 1
        GP0X = 206
    Case 2
        GP0X = 389
    Case 3
        GP0X = 390
    Case 4
        GP0X = 393
    Case 5
        GP0X = 426
    Case 6
        GP0X = 427
    Case 7
        GP0X = 426
    Case 8
        GP0X = 424
    Case 9
        GP0X = 423
    Case 10
        GP0X = 422
    Case 11
        GP0X = 421
    Case 12
        GP0X = 420
    Case 13
        GP0X = 419
    Case 14
        GP0X = 417
    Case 15
        GP0X = 416
    Case 16
        GP0X = 415
    Case 17
        GP0X = 414
    Case 18
        GP0X = 413
    Case 19
        GP0X = 412
    Case 20
        GP0X = 411
    Case 21
        GP0X = 409
    Case 22
        GP0X = 408
    Case 23
        GP0X = 407
    Case 24
        GP0X = 406
    Case 25
        GP0X = 405
    Case 26
        GP0X = 404
    Case 27
        GP0X = 403
    Case 28
        GP0X = 402
    Case 29
        GP0X = 401
    Case 30
        GP0X = 400
    Case 31
        GP0X = 399
    Case 32
        GP0X = 398
    Case 33
        GP0X = 396
    Case 34
        GP0X = 395
    Case 35
        GP0X = 394
    Case 36
        GP0X = 393
    Case 37
        GP0X = 392
    Case 38
        GP0X = 391
    Case 39
        GP0X = 390
    Case 40
        GP0X = 388
    Case 41
        GP0X = 387
    Case 42
        GP0X = 386
    Case 43
        GP0X = 385
    Case 44
        GP0X = 384
    Case 45
        GP0X = 383
    Case 46
        GP0X = 383
    Case 47
        GP0X = 384
    Case 48
        GP0X = 391
    Case 49
        GP0X = 402
    Case 50
        GP0X = 413
    Case 51
        GP0X = 424
    Case 52
        GP0X = 437
    Case 53
        GP0X = 448
    Case 54
        GP0X = 459
    Case 55
        GP0X = 470
    Case 56
        GP0X = 481
    Case 57
        GP0X = 492
    Case 58
        GP0X = 503
    Case 59
        GP0X = 514
    Case 60
        GP0X = 525
    Case 61
        GP0X = 536
    Case 62
        GP0X = 549
    Case 63
        GP0X = 560
    Case 64
        GP0X = 571
    Case 65
        GP0X = 582
    Case 66
        GP0X = 593
    Case 67
        GP0X = 594
    Case 68
        GP0X = 594
    Case 69
        GP0X = 593
    Case 70
        GP0X = 591
    Case 71
        GP0X = 590
    Case 72
        GP0X = 587
    Case 73
        GP0X = 582
    Case 74
        GP0X = 577
    Case 75
        GP0X = 572
    Case 76
        GP0X = 567
    Case 77
        GP0X = 562
    Case 78
        GP0X = 557
    Case 79
        GP0X = 552
    Case 80
        GP0X = 547
    Case 81
        GP0X = 542
    Case 82
        GP0X = 537
    Case 83
        GP0X = 532
    Case 84
        GP0X = 527
    Case 85
        GP0X = 522
    Case 86
        GP0X = 517
    Case 87
        GP0X = 512
    Case 88
        GP0X = 507
    Case 89
        GP0X = 502
    Case 90
        GP0X = 497
    Case 91
        GP0X = 492
    Case 92
        GP0X = 487
    Case 93
        GP0X = 482
    Case 94
        GP0X = 477
    Case 95
        GP0X = 474
    Case 96
        GP0X = 469
    Case 97
        GP0X = 464
    Case 98
        GP0X = 459
    Case 99
        GP0X = 454
    Case 100
        GP0X = 449
    Case 101
        GP0X = 444
    Case 102
        GP0X = 439
    Case 103
        GP0X = 434
    Case 104
        GP0X = 429
    Case 105
        GP0X = 424
    Case 106
        GP0X = 419
    Case 107
        GP0X = 414
    Case 108
        GP0X = 409
    Case 109
        GP0X = 404
    Case 110
        GP0X = 399
    Case 111
        GP0X = 394
    Case 112
        GP0X = 389
    Case 113
        GP0X = 384
    Case 114
        GP0X = 383
    Case 115
        GP0X = 384
    Case 116
        GP0X = 385
    Case 117
        GP0X = 386
    Case 118
        GP0X = 387
    Case 119
        GP0X = 388
    Case 120
        GP0X = 389
    Case 121
        GP0X = 390
    Case 122
        GP0X = 391
    Case 123
        GP0X = 392
    Case 124
        GP0X = 393
    Case 125
        GP0X = 394
    Case 126
        GP0X = 395
    Case 127
        GP0X = 396
    Case 128
        GP0X = 397
    Case 129
        GP0X = 398
    Case 130
        GP0X = 399
    Case 131
        GP0X = 400
    Case 132
        GP0X = 401
    Case 133
        GP0X = 402
    Case 134
        GP0X = 403
    Case 135
        GP0X = 404
    Case 136
        GP0X = 405
    Case 137
        GP0X = 406
    Case 138
        GP0X = 407
    Case 139
        GP0X = 408
    Case 140
        GP0X = 409
    Case 141
        GP0X = 410
    Case 142
        GP0X = 411
    Case 143
        GP0X = 412
    Case 144
        GP0X = 422
    Case 145
        GP0X = 423
    Case 146
        GP0X = 424
    Case 147
        GP0X = 425
    Case 148
        GP0X = 426
    Case 149
        GP0X = 427
    Case 150
        GP0X = 428
    Case 151
        GP0X = 429
    Case 152
        GP0X = 430
    Case 153
        GP0X = 431
    Case 154
        GP0X = 431
    Case 155
        GP0X = 432
    Case 156
        GP0X = 432
    Case 157
        GP0X = 431
    Case 158
        GP0X = 165
    Case 159
        GP0X = 164
    Case 160
        GP0X = 165
    Case 161
        GP0X = 166
    Case 162
        GP0X = 167
    Case 163
        GP0X = 168
    Case 164
        GP0X = 169
    Case 165
        GP0X = 170
    Case 166
        GP0X = 171
    Case 167
        GP0X = 172
    Case 168
        GP0X = 173
    Case 169
        GP0X = 174
    Case 170
        GP0X = 175
    Case 171
        GP0X = 176
    Case 172
        GP0X = 177
    Case 173
        GP0X = 178
    Case 174
        GP0X = 179
    Case 175
        GP0X = 180
    Case 176
        GP0X = 181
    Case 177
        GP0X = 182
    Case 178
        GP0X = 183
    Case 179
        GP0X = 184
    Case 180
        GP0X = 194
    Case 181
        GP0X = 195
    Case 182
        GP0X = 196
    Case 183
        GP0X = 197
    Case 184
        GP0X = 198
    Case 185
        GP0X = 199
    Case 186
        GP0X = 200
    Case 187
        GP0X = 201
    Case 188
        GP0X = 202
    Case 189
        GP0X = 203
    Case 190
        GP0X = 204
    Case 191
        GP0X = 205
    Case 192
        GP0X = 206
    Case 193
        GP0X = 207
    Case 194
        GP0X = 208
    Case 195
        GP0X = 209
    Case 196
        GP0X = 210
    Case 197
        GP0X = 211
    Case 198
        GP0X = 212
    Case 199
        GP0X = 213
    Case 200
        GP0X = 213
    Case 201
        GP0X = 212
    Case 202
        GP0X = 211
    Case 203
        GP0X = 209
    Case 204
        GP0X = 208
    Case 205
        GP0X = 205
    Case 206
        GP0X = 200
    Case 207
        GP0X = 195
    Case 208
        GP0X = 190
    Case 209
        GP0X = 185
    Case 210
        GP0X = 180
    Case 211
        GP0X = 175
    Case 212
        GP0X = 170
    Case 213
        GP0X = 165
    Case 214
        GP0X = 160
    Case 215
        GP0X = 155
    Case 216
        GP0X = 150
    Case 217
        GP0X = 145
    Case 218
        GP0X = 140
    Case 219
        GP0X = 135
    Case 220
        GP0X = 130
    Case 221
        GP0X = 125
    Case 222
        GP0X = 120
    Case 223
        GP0X = 117
    Case 224
        GP0X = 112
    Case 225
        GP0X = 107
    Case 226
        GP0X = 102
    Case 227
        GP0X = 97
    Case 228
        GP0X = 92
    Case 229
        GP0X = 87
    Case 230
        GP0X = 82
    Case 231
        GP0X = 77
    Case 232
        GP0X = 72
    Case 233
        GP0X = 67
    Case 234
        GP0X = 62
    Case 235
        GP0X = 57
    Case 236
        GP0X = 52
    Case 237
        GP0X = 47
    Case 238
        GP0X = 42
    Case 239
        GP0X = 37
    Case 240
        GP0X = 32
    Case 241
        GP0X = 27
    Case 242
        GP0X = 22
    Case 243
        GP0X = 17
    Case 244
        GP0X = 12
    Case 245
        GP0X = 7
    Case 246
        GP0X = 2
    Case 247
        GP0X = 1
    Case 248
        GP0X = 0
    Case 249
        GP0X = 0
    Case 250
        GP0X = 2
    Case 251
        GP0X = 3
    Case 252
        GP0X = 12
    Case 253
        GP0X = 23
    Case 254
        GP0X = 34
    Case 255
        GP0X = 45
    Case 256
        GP0X = 58
    Case 257
        GP0X = 69
    Case 258
        GP0X = 80
    Case 259
        GP0X = 91
    Case 260
        GP0X = 102
    Case 261
        GP0X = 113
    Case 262
        GP0X = 124
    Case 263
        GP0X = 135
    Case 264
        GP0X = 146
    Case 265
        GP0X = 157
    Case 266
        GP0X = 170
    Case 267
        GP0X = 181
    Case 268
        GP0X = 192
    Case 269
        GP0X = 203
    Case 270
        GP0X = 212
    Case 271
        GP0X = 213
    Case 272
        GP0X = 211
    Case 273
        GP0X = 210
    Case 274
        GP0X = 209
    Case 275
        GP0X = 208
    Case 276
        GP0X = 206
    Case 277
        GP0X = 205
    Case 278
        GP0X = 204
    Case 279
        GP0X = 203
    Case 280
        GP0X = 202
    Case 281
        GP0X = 201
    Case 282
        GP0X = 200
    Case 283
        GP0X = 198
    Case 284
        GP0X = 197
    Case 285
        GP0X = 196
    Case 286
        GP0X = 195
    Case 287
        GP0X = 194
    Case 288
        GP0X = 193
    Case 289
        GP0X = 192
    Case 290
        GP0X = 191
    Case 291
        GP0X = 190
    Case 292
        GP0X = 189
    Case 293
        GP0X = 188
    Case 294
        GP0X = 187
    Case 295
        GP0X = 185
    Case 296
        GP0X = 184
    Case 297
        GP0X = 183
    Case 298
        GP0X = 182
    Case 299
        GP0X = 181
    Case 300
        GP0X = 180
    Case 301
        GP0X = 179
    Case 302
        GP0X = 177
    Case 303
        GP0X = 176
    Case 304
        GP0X = 175
    Case 305
        GP0X = 174
    Case 306
        GP0X = 172
    Case 307
        GP0X = 171
    Case 308
        GP0X = 170
    Case 309
        GP0X = 169
    End Select
End Function
Private Function GP0Y(Number As Integer) As Integer
    Select Case Number
    Case 0
        GP0Y = 1
    Case 1
        GP0Y = 2
    Case 2
        GP0Y = 2
    Case 3
        GP0Y = 1
    Case 4
        GP0Y = 2
    Case 5
        GP0Y = 2
    Case 6
        GP0Y = 3
    Case 7
        GP0Y = 4
    Case 8
        GP0Y = 15
    Case 9
        GP0Y = 16
    Case 10
        GP0Y = 21
    Case 11
        GP0Y = 24
    Case 12
        GP0Y = 29
    Case 13
        GP0Y = 32
    Case 14
        GP0Y = 40
    Case 15
        GP0Y = 43
    Case 16
        GP0Y = 48
    Case 17
        GP0Y = 51
    Case 18
        GP0Y = 54
    Case 19
        GP0Y = 59
    Case 20
        GP0Y = 62
    Case 21
        GP0Y = 70
    Case 22
        GP0Y = 73
    Case 23
        GP0Y = 78
    Case 24
        GP0Y = 81
    Case 25
        GP0Y = 84
    Case 26
        GP0Y = 89
    Case 27
        GP0Y = 92
    Case 28
        GP0Y = 97
    Case 29
        GP0Y = 100
    Case 30
        GP0Y = 103
    Case 31
        GP0Y = 108
    Case 32
        GP0Y = 111
    Case 33
        GP0Y = 119
    Case 34
        GP0Y = 122
    Case 35
        GP0Y = 127
    Case 36
        GP0Y = 130
    Case 37
        GP0Y = 133
    Case 38
        GP0Y = 138
    Case 39
        GP0Y = 141
    Case 40
        GP0Y = 149
    Case 41
        GP0Y = 152
    Case 42
        GP0Y = 157
    Case 43
        GP0Y = 160
    Case 44
        GP0Y = 165
    Case 45
        GP0Y = 166
    Case 46
        GP0Y = 167
    Case 47
        GP0Y = 168
    Case 48
        GP0Y = 165
    Case 49
        GP0Y = 160
    Case 50
        GP0Y = 155
    Case 51
        GP0Y = 150
    Case 52
        GP0Y = 144
    Case 53
        GP0Y = 139
    Case 54
        GP0Y = 134
    Case 55
        GP0Y = 129
    Case 56
        GP0Y = 124
    Case 57
        GP0Y = 119
    Case 58
        GP0Y = 114
    Case 59
        GP0Y = 109
    Case 60
        GP0Y = 104
    Case 61
        GP0Y = 99
    Case 62
        GP0Y = 93
    Case 63
        GP0Y = 88
    Case 64
        GP0Y = 83
    Case 65
        GP0Y = 78
    Case 66
        GP0Y = 73
    Case 67
        GP0Y = 73
    Case 68
        GP0Y = 389
    Case 69
        GP0Y = 390
    Case 70
        GP0Y = 390
    Case 71
        GP0Y = 389
    Case 72
        GP0Y = 388
    Case 73
        GP0Y = 386
    Case 74
        GP0Y = 384
    Case 75
        GP0Y = 382
    Case 76
        GP0Y = 380
    Case 77
        GP0Y = 378
    Case 78
        GP0Y = 376
    Case 79
        GP0Y = 374
    Case 80
        GP0Y = 372
    Case 81
        GP0Y = 370
    Case 82
        GP0Y = 368
    Case 83
        GP0Y = 366
    Case 84
        GP0Y = 364
    Case 85
        GP0Y = 362
    Case 86
        GP0Y = 360
    Case 87
        GP0Y = 358
    Case 88
        GP0Y = 356
    Case 89
        GP0Y = 354
    Case 90
        GP0Y = 352
    Case 91
        GP0Y = 350
    Case 92
        GP0Y = 348
    Case 93
        GP0Y = 346
    Case 94
        GP0Y = 344
    Case 95
        GP0Y = 343
    Case 96
        GP0Y = 341
    Case 97
        GP0Y = 339
    Case 98
        GP0Y = 337
    Case 99
        GP0Y = 335
    Case 100
        GP0Y = 333
    Case 101
        GP0Y = 331
    Case 102
        GP0Y = 329
    Case 103
        GP0Y = 327
    Case 104
        GP0Y = 325
    Case 105
        GP0Y = 323
    Case 106
        GP0Y = 321
    Case 107
        GP0Y = 319
    Case 108
        GP0Y = 317
    Case 109
        GP0Y = 315
    Case 110
        GP0Y = 313
    Case 111
        GP0Y = 311
    Case 112
        GP0Y = 309
    Case 113
        GP0Y = 307
    Case 114
        GP0Y = 308
    Case 115
        GP0Y = 315
    Case 116
        GP0Y = 320
    Case 117
        GP0Y = 327
    Case 118
        GP0Y = 332
    Case 119
        GP0Y = 339
    Case 120
        GP0Y = 344
    Case 121
        GP0Y = 351
    Case 122
        GP0Y = 356
    Case 123
        GP0Y = 363
    Case 124
        GP0Y = 370
    Case 125
        GP0Y = 375
    Case 126
        GP0Y = 382
    Case 127
        GP0Y = 387
    Case 128
        GP0Y = 394
    Case 129
        GP0Y = 399
    Case 130
        GP0Y = 406
    Case 131
        GP0Y = 411
    Case 132
        GP0Y = 418
    Case 133
        GP0Y = 423
    Case 134
        GP0Y = 432
    Case 135
        GP0Y = 435
    Case 136
        GP0Y = 444
    Case 137
        GP0Y = 447
    Case 138
        GP0Y = 456
    Case 139
        GP0Y = 459
    Case 140
        GP0Y = 468
    Case 141
        GP0Y = 471
    Case 142
        GP0Y = 480
    Case 143
        GP0Y = 483
    Case 144
        GP0Y = 548
    Case 145
        GP0Y = 551
    Case 146
        GP0Y = 560
    Case 147
        GP0Y = 563
    Case 148
        GP0Y = 572
    Case 149
        GP0Y = 575
    Case 150
        GP0Y = 584
    Case 151
        GP0Y = 587
    Case 152
        GP0Y = 596
    Case 153
        GP0Y = 599
    Case 154
        GP0Y = 603
    Case 155
        GP0Y = 604
    Case 156
        GP0Y = 607
    Case 157
        GP0Y = 608
    Case 158
        GP0Y = 608
    Case 159
        GP0Y = 607
    Case 160
        GP0Y = 600
    Case 161
        GP0Y = 595
    Case 162
        GP0Y = 588
    Case 163
        GP0Y = 583
    Case 164
        GP0Y = 576
    Case 165
        GP0Y = 571
    Case 166
        GP0Y = 564
    Case 167
        GP0Y = 559
    Case 168
        GP0Y = 552
    Case 169
        GP0Y = 547
    Case 170
        GP0Y = 538
    Case 171
        GP0Y = 535
    Case 172
        GP0Y = 526
    Case 173
        GP0Y = 523
    Case 174
        GP0Y = 514
    Case 175
        GP0Y = 511
    Case 176
        GP0Y = 502
    Case 177
        GP0Y = 499
    Case 178
        GP0Y = 490
    Case 179
        GP0Y = 487
    Case 180
        GP0Y = 422
    Case 181
        GP0Y = 419
    Case 182
        GP0Y = 410
    Case 183
        GP0Y = 407
    Case 184
        GP0Y = 398
    Case 185
        GP0Y = 395
    Case 186
        GP0Y = 386
    Case 187
        GP0Y = 383
    Case 188
        GP0Y = 374
    Case 189
        GP0Y = 371
    Case 190
        GP0Y = 362
    Case 191
        GP0Y = 357
    Case 192
        GP0Y = 350
    Case 193
        GP0Y = 345
    Case 194
        GP0Y = 338
    Case 195
        GP0Y = 333
    Case 196
        GP0Y = 326
    Case 197
        GP0Y = 321
    Case 198
        GP0Y = 314
    Case 199
        GP0Y = 309
    Case 200
        GP0Y = 307
    Case 201
        GP0Y = 306
    Case 202
        GP0Y = 307
    Case 203
        GP0Y = 307
    Case 204
        GP0Y = 308
    Case 205
        GP0Y = 309
    Case 206
        GP0Y = 311
    Case 207
        GP0Y = 313
    Case 208
        GP0Y = 315
    Case 209
        GP0Y = 317
    Case 210
        GP0Y = 319
    Case 211
        GP0Y = 321
    Case 212
        GP0Y = 323
    Case 213
        GP0Y = 325
    Case 214
        GP0Y = 327
    Case 215
        GP0Y = 329
    Case 216
        GP0Y = 331
    Case 217
        GP0Y = 333
    Case 218
        GP0Y = 335
    Case 219
        GP0Y = 337
    Case 220
        GP0Y = 339
    Case 221
        GP0Y = 341
    Case 222
        GP0Y = 343
    Case 223
        GP0Y = 344
    Case 224
        GP0Y = 346
    Case 225
        GP0Y = 348
    Case 226
        GP0Y = 350
    Case 227
        GP0Y = 352
    Case 228
        GP0Y = 354
    Case 229
        GP0Y = 356
    Case 230
        GP0Y = 358
    Case 231
        GP0Y = 360
    Case 232
        GP0Y = 362
    Case 233
        GP0Y = 364
    Case 234
        GP0Y = 366
    Case 235
        GP0Y = 368
    Case 236
        GP0Y = 370
    Case 237
        GP0Y = 372
    Case 238
        GP0Y = 374
    Case 239
        GP0Y = 376
    Case 240
        GP0Y = 378
    Case 241
        GP0Y = 380
    Case 242
        GP0Y = 382
    Case 243
        GP0Y = 384
    Case 244
        GP0Y = 386
    Case 245
        GP0Y = 388
    Case 246
        GP0Y = 390
    Case 247
        GP0Y = 390
    Case 248
        GP0Y = 387
    Case 249
        GP0Y = 73
    Case 250
        GP0Y = 73
    Case 251
        GP0Y = 74
    Case 252
        GP0Y = 78
    Case 253
        GP0Y = 83
    Case 254
        GP0Y = 88
    Case 255
        GP0Y = 93
    Case 256
        GP0Y = 99
    Case 257
        GP0Y = 104
    Case 258
        GP0Y = 109
    Case 259
        GP0Y = 114
    Case 260
        GP0Y = 119
    Case 261
        GP0Y = 124
    Case 262
        GP0Y = 129
    Case 263
        GP0Y = 134
    Case 264
        GP0Y = 139
    Case 265
        GP0Y = 144
    Case 266
        GP0Y = 150
    Case 267
        GP0Y = 155
    Case 268
        GP0Y = 160
    Case 269
        GP0Y = 165
    Case 270
        GP0Y = 169
    Case 271
        GP0Y = 168
    Case 272
        GP0Y = 160
    Case 273
        GP0Y = 157
    Case 274
        GP0Y = 152
    Case 275
        GP0Y = 149
    Case 276
        GP0Y = 141
    Case 277
        GP0Y = 138
    Case 278
        GP0Y = 133
    Case 279
        GP0Y = 130
    Case 280
        GP0Y = 127
    Case 281
        GP0Y = 122
    Case 282
        GP0Y = 119
    Case 283
        GP0Y = 111
    Case 284
        GP0Y = 108
    Case 285
        GP0Y = 103
    Case 286
        GP0Y = 100
    Case 287
        GP0Y = 97
    Case 288
        GP0Y = 92
    Case 289
        GP0Y = 89
    Case 290
        GP0Y = 84
    Case 291
        GP0Y = 81
    Case 292
        GP0Y = 78
    Case 293
        GP0Y = 73
    Case 294
        GP0Y = 70
    Case 295
        GP0Y = 62
    Case 296
        GP0Y = 59
    Case 297
        GP0Y = 54
    Case 298
        GP0Y = 51
    Case 299
        GP0Y = 48
    Case 300
        GP0Y = 43
    Case 301
        GP0Y = 40
    Case 302
        GP0Y = 32
    Case 303
        GP0Y = 29
    Case 304
        GP0Y = 24
    Case 305
        GP0Y = 21
    Case 306
        GP0Y = 13
    Case 307
        GP0Y = 10
    Case 308
        GP0Y = 5
    Case 309
        GP0Y = 2
    End Select
End Function

Private Sub cmdbiblebase_Click()
frmBibleBase.Show       'Displays the Bible base form
End Sub

Private Sub Button_Click(Index As Integer)

Dim reply As Integer

Select Case Index
    Case 0
        frmDiary.Show
        Unload Me
    Case 1
        frmPrayers.Show
        Unload Me
    Case 2
        frmChurchYear.Show
        Unload Me
    Case 3
        frmMassSheet.Show
        Unload Me
    Case 4
        reply = MsgBox("Are you sure you want to quit?", vbYesNo, "The Genesis Project")
        If reply = 6 Then End
    Case 5
        frmBibleBase.Show
        Unload Me
    Case 6
        frmDeathRecords.Show
        Unload Me
    Case 7
        frmhelp.Show
        Unload Me
End Select

End Sub

Private Sub Button_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    drawbuttons
    ShapedForm.Button(Index).Picture = ShapedForm.Button(8).Picture
Select Case Index
Case 0
    lblChosenOption.Caption = "Store all of youre upcoming events, with detals and times, reminder function will be in future version. Has search function."
Case 1
    lblChosenOption.Caption = "This will allow you to store prayers and print them at will. Allows you to search through them with a single click."
Case 2
    lblChosenOption.Caption = "Requires internet, will allow you to navigate to the roman catholic online rescorse to review readings and nessasery information."
Case 3
    lblChosenOption.Caption = "Allows you to follow a wizard that will guide you through producing a mass sheet with custom readings, major prayers and layout."
Case 4
    lblChosenOption.Caption = "Quit the genesis project"
Case 5
    lblChosenOption.Caption = "Allows you to browse the bible and review passages, the selected passage can be read and tweaked manually before printing."
Case 6
    lblChosenOption.Caption = "Store death records in a database for reference"
Case 7
    lblChosenOption.Caption = "How to enter a bible address"
End Select
End Sub

Private Sub Form_Click()
    End
End Sub

Private Sub Form_Load()
    Dim nRet As Long
    nRet = SetWindowRgn(Me.hWnd, CreateFormRegion(1, 1, 0, 0), True)
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hWnd, &HA1, 2, 0&
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
drawbuttons
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DeleteObject ResultRegion
End Sub

Sub drawbuttons()
Dim counter As Integer
Do
    Button(counter).Picture = Button(9)
    counter = counter + 1
    lblChosenOption.Caption = "Please Select"
Loop Until counter = 8

End Sub

