VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Thumbnail Demo"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11460
   Icon            =   "dsThumb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   11460
   StartUpPosition =   2  'CenterScreen
   Begin vb6projectProject1.dsThumb dsThumb1 
      Align           =   4  'Align Right
      Height          =   4680
      Left            =   3840
      TabIndex        =   9
      Top             =   0
      Width           =   7620
      _ExtentX        =   13441
      _ExtentY        =   8255
      Cols            =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Rows            =   3
      ThumbnailSize   =   1
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Fore Color Text"
      Height          =   495
      Left            =   1800
      TabIndex        =   8
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Back Color Text"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   4800
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Back Color Select"
      Height          =   495
      Left            =   1800
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Back Color Frame"
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   495
      Left            =   1800
      TabIndex        =   2
      Top             =   5400
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Thumb"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "Thumb Count"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   1320
      TabIndex        =   3
      Top             =   2640
      Width           =   1860
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Height          =   2415
      Left            =   120
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   375
      Left            =   3480
      TabIndex        =   1
      Top             =   5520
      Width           =   7260
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command3_Click()
 dsThumb1.BackColorFrame = vbYellow
End Sub

Private Sub Command4_Click()
 dsThumb1.BackColorSel = vbRed
End Sub

Private Sub Command5_Click()
 dsThumb1.BackColorText = vbGreen
End Sub

Private Sub Command6_Click()
 dsThumb1.ForeColorText = vbBlue
End Sub

Private Sub Command1_Click()
 dsThumb1.ThumbPath = App.Path & "\Thumbs"
 dsThumb1.SetFocus
 Label1.Caption = dsThumb1.GetFileName
 Label2.Caption = dsThumb1.ThumbCount
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub dsThumb1_Click()
 Label1.Caption = dsThumb1.GetFileName
 Set Image1.Picture = dsThumb1.GetPicture
End Sub

Private Sub dsThumb1_KeyDown(KeyCode As Integer, Shift As Integer)
 Label1.Caption = dsThumb1.GetFileName
 Set Image1.Picture = dsThumb1.GetPicture
End Sub

