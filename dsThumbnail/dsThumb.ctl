VERSION 5.00
Begin VB.UserControl dsThumb 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4980
   ControlContainer=   -1  'True
   EditAtDesignTime=   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   356
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   332
   ToolboxBitmap   =   "dsThumb.ctx":0000
   Begin VB.PictureBox picFrame 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5340
      Left            =   0
      ScaleHeight     =   356
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   315
      TabIndex        =   2
      Top             =   0
      Width           =   4725
      Begin VB.FileListBox filHidden 
         Height          =   675
         Left            =   600
         Pattern         =   "*.bmp;*.dib;*.rle;*.gif;*.jpg;*.wmf;*.emf;*.ico;*.cur;*.png"
         TabIndex        =   5
         Top             =   4560
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.PictureBox picSlide 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1875
         Left            =   0
         ScaleHeight     =   125
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   165
         TabIndex        =   3
         Top             =   0
         Visible         =   0   'False
         Width           =   2475
         Begin VB.PictureBox picBack 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            ForeColor       =   &H80000008&
            Height          =   1560
            Index           =   0
            Left            =   120
            ScaleHeight     =   102
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   142
            TabIndex        =   6
            Top             =   120
            Width           =   2160
            Begin VB.PictureBox picText 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00C0FFC0&
               ForeColor       =   &H80000008&
               Height          =   315
               Index           =   0
               Left            =   15
               ScaleHeight     =   19
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   138
               TabIndex        =   7
               Top             =   1200
               Width           =   2100
            End
            Begin VB.PictureBox picThumb 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00FFFFFF&
               ForeColor       =   &H80000008&
               Height          =   1200
               Index           =   0
               Left            =   15
               ScaleHeight     =   78
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   138
               TabIndex        =   0
               Top             =   15
               Width           =   2100
            End
         End
         Begin VB.PictureBox picFocus 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFF80&
            ForeColor       =   &H80000008&
            Height          =   1680
            Left            =   60
            ScaleHeight     =   110
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   150
            TabIndex        =   4
            Top             =   60
            Visible         =   0   'False
            Width           =   2280
         End
      End
      Begin VB.Timer FocusTimer 
         Left            =   120
         Top             =   4800
      End
   End
   Begin VB.VScrollBar VScroll 
      Enabled         =   0   'False
      Height          =   5340
      Left            =   4710
      Max             =   115
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "dsThumb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'GDI Plus API
Private Declare Sub GdiplusShutdown Lib "Gdiplus" (ByVal Token As Long)
Private Declare Function GdipCreateFromHDC Lib "Gdiplus" (ByVal hDC As Long, graphics As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "Gdiplus" (ByVal graphics As Long) As Long
Private Declare Function GdipLoadImageFromFile Lib "Gdiplus" (ByVal FileName As String, Image As Long) As Long
Private Declare Function GdipGetImageWidth Lib "Gdiplus" (ByVal Image As Long, Width As Long) As Long
Private Declare Function GdipGetImageHeight Lib "Gdiplus" (ByVal Image As Long, Height As Long) As Long
Private Declare Function GdipDisposeImage Lib "Gdiplus" (ByVal Image As Long) As Long
Private Declare Function GdiplusStartup Lib "Gdiplus" (Token As Long, _
    InputBuffer As GdiplusStartupInput, Optional ByVal OutputBuffer As Long = 0) As Long
'For thumbnail
Private Declare Function GdipGetImageThumbnail Lib "Gdiplus" (ByVal Image As Long, _
    ByVal ThumbWidth As Long, ByVal ThumbHeight As Long, ThumbImage As Long, _
    Optional ByVal Callback As Long = 0, Optional ByVal CallbackData As Long = 0) As Long
'Draw image
Private Declare Function GdipDrawImageRect Lib "Gdiplus" (ByVal graphics As Long, _
    ByVal Image As Long, ByVal X As Single, ByVal Y As Single, _
    ByVal Width As Single, ByVal Height As Single) As Long

' API Declarations
' ==================================
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function PeekMessage Lib "user32" Alias "PeekMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long, ByVal wRemoveMsg As Long) As Long
Private Declare Function WaitMessage Lib "user32" () As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function GetTickCount& Lib "kernel32" ()
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal x1 As Long, ByVal Y1 As Long, _
    ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, _
    ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

Private Type GdiplusStartupInput
 GdiplusVersion           As Long    'Must be 1 for Gdi+ v1.0, the current version as of this writing.
 DebugEventCallback       As Long    'Ignored on free builds
 SuppressBackgroundThread As Long    'FALSE unless you're prepared to call the hook/unhook functions properly
 SuppressExternalCodecs   As Long    'FALSE unless you want Gdi+ only to use its internal image codecs.
End Type

Private Type IMGRECT
 Left                    As Long
 Top                     As Long
 Width                   As Long
 Height                  As Long
End Type

Private Type RECT
 Left                    As Long
 Top                     As Long
 Right                   As Long
 Bottom                  As Long
End Type

Private Type TYPEGDITHUMBNAILSIMAGE
 strFile                 As String
 mDC                     As Long
 hImage                  As Long
 cRECT                   As IMGRECT
 orgWidth                As Long
 orgHeight               As Long
 bSelected               As Boolean
End Type

Private Type POINTAPI
 X As Long
 Y As Long
End Type

Private Type msg
 hwnd As Long
 Message As Long
 wParam As Long
 lParam As Long
 time As Long
 pt As POINTAPI
End Type

Private Type CurrentControlType
 Name As String
 Index As Long
End Type

Public Enum BorderStyleEnum
 [None]
 [Fixed Single]
End Enum

Public Enum ThumbnailSizeEnum
 [Tinny]
 [Small]
 [Medium]
 [Large]
 [Extra Large]
End Enum

Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)

Private Const WM_MOUSEWHEEL = 522
Private Const SM_CYVSCROLL = 20
Private Const PM_REMOVE = &H1

Private Token As Long
Private mItem() As TYPEGDITHUMBNAILSIMAGE
Private lImg() As Long
Private nCount As Long
Private flName() As String

Private tWidth As Long
Private tHeight As Long
Private bCancel As Boolean
Dim curThumb As Long
Dim scrollByArrow As Boolean

'Default Property Values:
Const m_def_BackColorFrame = &H808080
Const m_def_BackColorSel = &HFFFF80
Const m_def_BackColorText = &HE0E0E0
Const m_def_BorderStyle = 1
Const m_def_Cols = 2
Const m_def_Enabled = True
Const m_def_ForeColorText = &H80000008
Const m_def_Rows = 2
Const m_def_ThumbPath = "C:\"
Const m_def_ThumbnailSize = 2

'Property Variables:
Dim m_BackColorFrame As OLE_COLOR
Dim m_BackColorSel As OLE_COLOR
Dim m_BackColorText As OLE_COLOR
Dim m_BorderStyle As BorderStyleEnum
Dim m_Cols As Long
Dim m_Enabled As Boolean
Dim m_Font As StdFont
Dim m_ForeColorText As OLE_COLOR
Dim m_Rows As Long
Dim m_ThumbCount As Long
Dim m_ThumbPath As String
Dim m_ThumbnailSize As ThumbnailSizeEnum

Private Sub UserControl_Initialize()
   
 Dim mInput   As GdiplusStartupInput
 mInput.GdiplusVersion = 1
   
 If GdiplusStartup(Token, mInput) <> 0 Then          ' Unable to load GDI+
  MsgBox "Error loading GDIPlus!", vbExclamation + vbOKOnly
  Unload Me
 End If
    
 VScroll.Value = 0
 
End Sub

Private Sub UserControl_InitProperties()
 m_BackColorFrame = m_def_BackColorFrame
 m_BackColorSel = m_def_BackColorSel
 m_BackColorText = m_def_BackColorText
 m_BorderStyle = m_def_BorderStyle
 m_Cols = m_def_Cols
 m_Enabled = m_def_Enabled
 Set m_Font = Ambient.Font
 m_ForeColorText = m_def_ForeColorText
 m_Rows = m_def_Rows
 m_ThumbPath = m_def_ThumbPath
 m_ThumbnailSize = m_def_ThumbnailSize
 Call SetFrame(m_ThumbnailSize)
End Sub

Private Sub UserControl_Resize()
 
 On Error Resume Next
 
 With UserControl
  
  picFocus.Height = tHeight + 8
  picFocus.Width = tWidth + 8
  
  picBack(0).Height = tHeight
  picBack(0).Width = tWidth
  
  picThumb(0).Height = tHeight - 24
  picThumb(0).Width = tWidth - 4
  
  picText(0).Top = picThumb(0).Height
  picText(0).Height = 21
  picText(0).Width = tWidth - 4
  
  .Height = ((m_Rows * (picFocus.Height) + 8) + 4) * 15
  .Width = ((m_Cols * (picFocus.Width) + 8) + VScroll.Width + 3) * 15
  
  picFrame.Height = (m_Rows * (picFocus.Height) + 8) + 4
  picFrame.Width = (m_Cols * (tWidth + 8) + 8)
  
  VScroll.Height = picFrame.Height - 4
  VScroll.Left = picFrame.Width
  VScroll.Enabled = False
  
 End With
 
End Sub

Private Sub UserControl_Paint()
 If Not Ambient.UserMode Then UserControl_Resize
End Sub

Private Sub UserControl_Show()
 UserControl_Resize
 FocusTimer.Enabled = True
 ProcessMessages
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

 UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
 m_BackColorFrame = PropBag.ReadProperty("BackColorFrame", m_def_BackColorFrame)
 m_BackColorSel = PropBag.ReadProperty("BackColorSel", m_def_BackColorSel)
 m_BackColorText = PropBag.ReadProperty("BackColorText", m_def_BackColorText)
 m_BorderStyle = PropBag.ReadProperty("BorderStyle", m_def_BorderStyle)
 m_Cols = PropBag.ReadProperty("Cols", m_def_Cols)
 m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
 Set m_Font = PropBag.ReadProperty("Font", Ambient.Font)
 m_ForeColorText = PropBag.ReadProperty("ForeColorText", m_def_ForeColorText)
 m_Rows = PropBag.ReadProperty("Rows", m_def_Rows)
 m_ThumbPath = PropBag.ReadProperty("ThumbPath", m_def_ThumbPath)
 m_ThumbnailSize = PropBag.ReadProperty("ThumbnailSize", m_def_ThumbnailSize)
 picFrame.BackColor = m_BackColorFrame
 picSlide.BackColor = m_BackColorFrame
 picFocus.BackColor = m_BackColorSel
 picText(0).BackColor = m_BackColorText
 picText(0).ForeColor = m_ForeColorText
 Set picText(0).Font = m_Font
 Call SetFrame(m_ThumbnailSize)
 
 If Ambient.UserMode Then
  Call CreateThumbs
 End If
 
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
 Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 1)
 Call PropBag.WriteProperty("BackColorFrame", m_BackColorFrame, m_def_BackColorFrame)
 Call PropBag.WriteProperty("BackColorSel", m_BackColorSel, m_def_BackColorSel)
 Call PropBag.WriteProperty("BackColorText", m_BackColorText, m_def_BackColorText)
 Call PropBag.WriteProperty("Cols", m_Cols, m_def_Cols)
 Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
 Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
 Call PropBag.WriteProperty("ForeColorText", m_ForeColorText, m_def_ForeColorText)
 Call PropBag.WriteProperty("Rows", m_Rows, m_def_Rows)
 Call PropBag.WriteProperty("ThumbPath", m_ThumbPath, m_def_ThumbPath)
 Call PropBag.WriteProperty("ThumbnailSize", m_ThumbnailSize, m_def_ThumbnailSize)
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
 bCancel = True              'IDE SAFE method, if go to debug mode,
 FocusTimer.Enabled = False  'Stop ProcessMessages
End Sub

Private Sub UserControl_Terminate()
 Call RemoveThumbs
 bCancel = True
 FocusTimer = False
 GdiplusShutdown Token
End Sub

'User control properties
' ==================================
Public Property Get BackColorFrame() As OLE_COLOR
 BackColorFrame = m_BackColorFrame
End Property

Public Property Let BackColorFrame(ByVal New_BackColorFrame As OLE_COLOR)
 m_BackColorFrame = New_BackColorFrame
 PropertyChanged "BackColorFrame"
 picFrame.BackColor = New_BackColorFrame
 picSlide.BackColor = New_BackColorFrame
End Property

Public Property Get BackColorSel() As OLE_COLOR
 BackColorSel = m_BackColorSel
End Property

Public Property Let BackColorSel(ByVal New_BackColorSel As OLE_COLOR)
 m_BackColorSel = New_BackColorSel
 PropertyChanged "BackColorSel"
 picFocus.BackColor = New_BackColorSel
End Property

Public Property Get BackColorText() As OLE_COLOR
 BackColorText = m_BackColorText
End Property

Public Property Let BackColorText(ByVal New_BackColorText As OLE_COLOR)
 m_BackColorText = New_BackColorText
 PropertyChanged "BackColorText"
 If Ambient.UserMode Then 'Run Mode
  Call ChangeTextProp
 End If
End Property

Public Property Get BorderStyle() As BorderStyleEnum
 BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BorderStyleEnum)
 UserControl.BorderStyle() = New_BorderStyle
 PropertyChanged "BorderStyle"
End Property

Public Property Get Cols() As Long
 Cols = m_Cols
End Property

Public Property Let Cols(ByVal New_Cols As Long)
 If Ambient.UserMode Then 'Run Mode
  MsgBox "Allowed in design time only.", vbCritical, App.Comments
  Exit Property
 Else 'Design Mode
  m_Cols = New_Cols
  PropertyChanged "Cols"
  UserControl_Resize
 End If
End Property

Public Property Get Enabled() As Boolean
 Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
 m_Enabled = New_Enabled
 PropertyChanged "Enabled"
End Property

Public Property Get Font() As Font
 Set Font = m_Font
End Property

Public Property Set Font(ByVal New_Font As Font)

 Set m_Font = New_Font
 UserControl.Font = New_Font
 Set picText(0).Font = New_Font
 PropertyChanged "Font"
 
 If Ambient.UserMode Then 'Run Mode
  Call ChangeTextProp
 End If
 
End Property

Public Property Get ForeColorText() As OLE_COLOR
 ForeColorText = m_ForeColorText
End Property

Public Property Let ForeColorText(ByVal New_ForeColorText As OLE_COLOR)
 m_ForeColorText = New_ForeColorText
 PropertyChanged "ForeColorText"
 If Ambient.UserMode Then 'Run Mode
  Call ChangeTextProp
 End If
End Property

Public Property Get Rows() As Long
 Rows = m_Rows
End Property

Public Property Let Rows(ByVal New_Rows As Long)
 If Ambient.UserMode Then 'Run Mode
  MsgBox "Allowed in design time only.", vbCritical, App.Comments
  Exit Property
 Else
  m_Rows = New_Rows
  PropertyChanged "Rows"
  UserControl_Resize
 End If
End Property

Public Property Get ThumbCount() As String
 ThumbCount = nCount
End Property

Public Property Get ThumbPath() As String
 ThumbPath = m_ThumbPath
End Property

Public Property Let ThumbPath(ByVal New_ThumbPath As String)
 m_ThumbPath = New_ThumbPath
 PropertyChanged "ThumbPath"
 If Ambient.UserMode Then
  Call CreateThumbs
 End If
End Property

Public Property Get ThumbnailSize() As ThumbnailSizeEnum
 ThumbnailSize = m_ThumbnailSize
End Property

Public Property Let ThumbnailSize(ByVal New_ThumbnailSize As ThumbnailSizeEnum)
 
 If Ambient.UserMode Then 'Run Mode
  MsgBox "Allowed in design time only.", vbCritical, App.Comments
  Exit Property
 Else
  m_ThumbnailSize = New_ThumbnailSize
  PropertyChanged "ThumbnailSize"
  Call SetFrame(New_ThumbnailSize)
  UserControl_Resize
 End If
 
End Property

Public Property Get hwnd() As Long
 hwnd = UserControl.hwnd
End Property

Public Function GetPicture() As StdPicture
 If nCount - 1 > 0 Then
  Set GetPicture = picThumb(curThumb).Image
  'Set GetPicture = LoadPicture(mItem(curThumb).strFile)
 Else
  MsgBox "No thumbnail available.", vbExclamation
  Exit Function
 End If
End Function

Public Function GetFileName() As String
 If nCount - 1 > 0 Then
  GetFileName = mItem(curThumb).strFile
 Else
  MsgBox "No thumbnail available.", vbExclamation
  Exit Function
 End If
End Function

Public Sub ShowAbout()
Attribute ShowAbout.VB_UserMemId = -552
 frmAbout.Show
End Sub

Private Sub VScroll_GotFocus()
 picThumb(curThumb).SetFocus
End Sub

Private Sub VScroll_Change()
 picSlide.Top = -VScroll.Value
 If scrollByArrow = False Then
  picThumb(curThumb).SetFocus
 End If
End Sub

Private Sub VScroll_Scroll()
 VScroll_Change
End Sub

Private Sub picSlide_Click()
 picThumb(curThumb).SetFocus
End Sub

Private Sub picFocus_Click()
 picThumb(curThumb).SetFocus
 RaiseEvent Click
End Sub

Private Sub picBack_Click(Index As Integer)
 picThumb(Index).SetFocus
 RaiseEvent Click
End Sub

Private Sub picText_Click(Index As Integer)
 picThumb(Index).SetFocus
 RaiseEvent Click
End Sub

Private Sub picThumb_GotFocus(Index As Integer)
 curThumb = Index
 picFocus.Move picBack(Index).Left - 4, picBack(Index).Top - 4
 picFocus.Visible = True
 picThumb(Index).SetFocus
End Sub

Private Sub picThumb_Click(Index As Integer)
 curThumb = Index
 picFocus.Move picBack(Index).Left - 4, picBack(Index).Top - 4
 picFocus.Visible = True
 picThumb(Index).SetFocus
 RaiseEvent Click
End Sub

Private Sub picThumb_DblClick(Index As Integer)
 curThumb = Index
 picFocus.Move picBack(Index).Left - 4, picBack(Index).Top - 4
 picFocus.Visible = True
 picThumb(Index).SetFocus
 RaiseEvent DblClick
End Sub

Private Sub picThumb_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
 
 Dim q As Integer
 Dim Result As Long
 Dim picTop As Long
 Dim remThumb As Long
  
 If KeyCode = vbKeyDown Then 'Down Arrow
  
  If Index + m_Cols < (nCount) Then 'Next Focus
   picThumb(Index + m_Cols).SetFocus
   
   If picSlide.Top = 0 Then
    picTop = 0
   Else
    picTop = Mid(picSlide.Top, 2, 10) / (tHeight + 8)
   End If
      
   If (Index) < ((m_Rows * m_Cols) + (m_Cols * picTop)) And (Index) >= ((m_Rows * m_Cols) + (m_Cols * (picTop - 1))) Then 'Scroll Down
    scrollByArrow = True
    VScroll.Value = VScroll.Value + (tHeight + 8)
   End If
   
  End If
  
 ElseIf KeyCode = vbKeyUp Then 'Up Arrow
  
  If Index >= m_Cols Then 'Next Focus
   picThumb(Index - m_Cols).SetFocus
   
   If picSlide.Top = 0 Then
    picTop = 0
   Else
    picTop = Mid(picSlide.Top, 2, 10) / (tHeight + 8)
   End If
   
   If (Index) >= (m_Cols * picTop) And (Index) < (m_Cols * (picTop + 1)) Then 'Scroll Up
    scrollByArrow = True
    VScroll.Value = VScroll.Value - (tHeight + 8)
   End If
   
  End If
  
 ElseIf KeyCode = vbKeyRight Then 'Right Arrow
  
  If Index < nCount - 1 Then
   If GetCellNo(Index) < m_Cols Then 'Next Focus
    picThumb(Index + 1).SetFocus
   End If
  End If
   
 ElseIf KeyCode = vbKeyLeft Then 'Left Arrow
 
  If GetCellNo(Index) > 1 Then
   picThumb(Index - 1).SetFocus
  End If
   
 ElseIf KeyCode = vbKeyHome And Shift = 2 Then 'Ctrl Home
  VScroll.Value = 0
  picThumb(0).SetFocus
   
 ElseIf KeyCode = vbKeyEnd And Shift = 2 Then 'Ctrl End
  VScroll.Value = VScroll.Max
  picThumb(picThumb.Count - 1).SetFocus
 
 ElseIf KeyCode = vbKeyHome Then 'Home
 
  If picSlide.Top = 0 Then
   picTop = 0
  Else
   picTop = Mid(picSlide.Top, 2, 10) / (tHeight + 8)
  End If
  
  Result = m_Cols * picTop
  picThumb(Result).SetFocus
  
 ElseIf KeyCode = vbKeyEnd Then 'End
  
  If picSlide.Top = 0 Then
   picTop = 0
  Else
   picTop = Mid(picSlide.Top, 2, 10) / (tHeight + 8)
  End If
   
  If (((m_Rows * m_Cols) + (m_Cols * picTop)) - 1) > nCount - 1 Then
   Result = nCount - 1
  Else
   Result = ((m_Rows * m_Cols) + (m_Cols * picTop)) - 1
  End If
  
  picThumb(Result).SetFocus
  
 ElseIf KeyCode = vbKeyPageDown Then 'Page Down
    
  Result = Int((Index + m_Cols) / m_Cols)
  Result = Result * (tHeight + 8)
  
  If Result > VScroll.Max Then
   VScroll.Value = VScroll.Max
  Else
   VScroll.Value = Result
  End If
  
  If (Index + (m_Rows * m_Cols)) > nCount - 1 Then
    
   For q = 0 To (GetCellNo(nCount - 1) - 1)
    If picBack((nCount - 1) - q).Left = picBack(Index).Left Then
     remThumb = (nCount - 1) - q
     Exit For
    Else
     remThumb = nCount - 1
    End If
   Next q
    
   picThumb(remThumb).SetFocus
  Else
   picThumb(Index + (m_Rows * m_Cols)).SetFocus
  End If
  
 ElseIf KeyCode = vbKeyPageUp Then 'Page Up
   
  Result = Int(((nCount - 1) - (Index)) / m_Cols) + 1
  Result = Result * (tHeight + 8)
     
  If Result > VScroll.Max Then
   VScroll.Value = VScroll.Min
  Else
   VScroll.Value = VScroll.Max - Result
  End If
      
  If (Index - (m_Rows * m_Cols)) < 0 Then
  
   For q = 3 To 0 Step -1
    If picBack(q).Left = picBack(Index).Left Then
     remThumb = q
     Exit For
    Else
     remThumb = 0
    End If
   Next q
  
   picThumb(remThumb).SetFocus
  Else
   picThumb(Index - (m_Rows * m_Cols)).SetFocus
  End If
   
 End If
 
 DoEvents
 RaiseEvent KeyDown(KeyCode, Shift)
 
End Sub

Private Sub picThumb_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
 RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub picThumb_LostFocus(Index As Integer)
 picFocus.Visible = False
End Sub

Private Sub CreateThumbs()
 
 On Error Resume Next
 Dim intCounter, LF, TP As Integer
 Dim strFile As String
 Dim q As Integer
 
 filHidden.Path = m_ThumbPath
 nCount = filHidden.ListCount
 ReDim mItem(nCount - 1)
 ReDim flName(nCount - 1)
 If nCount <= 0 Then Exit Sub
 
 For intCounter = 0 To nCount - 1
  mItem(intCounter).strFile = filHidden.Path & "\" & filHidden.List(intCounter)
 Next
 
 For intCounter = 0 To nCount - 1 'Getting File Name
  
  For q = Len(filHidden.List(intCounter)) To 0 Step -1
   strFile = Mid(filHidden.List(intCounter), 1, q)
   If picThumb(0).TextWidth(strFile) < picThumb(0).Width - 30 Then
    flName(intCounter) = strFile
    Exit For
   End If
  Next q
   
  If Len(strFile) < Len(filHidden.List(intCounter)) Then
   flName(intCounter) = flName(intCounter) + "..."
  End If
  
 Next
 
 LF = 8 'Set Default Left
 TP = 8  'Set Default Top
 intCounter = 0    'Counter
 picThumb(0).ToolTipText = mItem(0).strFile
 
 Screen.MousePointer = vbHourglass
 picSlide.Visible = False
 
 For intCounter = 1 To nCount - 1
     
  LF = LF + picFocus.Width
  If LF >= (tWidth * m_Cols) + ((m_Cols + 1) * 8) Then
   LF = 8
   TP = TP + picFocus.Height
  End If
  
  Load picBack(intCounter)
  Load picThumb(intCounter)
  Load picText(intCounter)
  
  picBack(intCounter).Left = LF
  picBack(intCounter).Top = TP
  
  picBack(intCounter).ToolTipText = mItem(intCounter).strFile
  picThumb(intCounter).ToolTipText = mItem(intCounter).strFile
  picText(intCounter).ToolTipText = mItem(intCounter).strFile
  picThumb(intCounter).TabIndex = intCounter
  
  Set picBack(intCounter).Container = picSlide
  Set picThumb(intCounter).Container = picBack(intCounter)
  Set picText(intCounter).Container = picBack(intCounter)
    
  picThumb(intCounter).Visible = True
  picBack(intCounter).Visible = True
  picText(intCounter).Visible = True
  
 Next intCounter
   
 picSlide.Width = m_Cols * (picFocus.Width) + 8
 picSlide.Height = Int(picThumb.Count / m_Cols) * (picFocus.Height) + 8
 
 If Int(picThumb.Count / m_Cols) < (picThumb.Count / m_Cols) Then
  picSlide.Height = picSlide.Height + (picFocus.Height)
 End If
  
 Call GenerateThumb
 Call Draw
  
 VScroll.Value = 0
 VScroll.Max = (picSlide.Height - picFrame.ScaleHeight) + 4
  
 If VScroll.Max < 0 Then
  VScroll.Max = 0
  VScroll.Enabled = False
 Else
  VScroll.Enabled = True
  VScroll.SmallChange = picFocus.Height
  VScroll.LargeChange = m_Rows * picFocus.Height
 End If
    
 picSlide.Visible = True
 Screen.MousePointer = vbNormal
 
End Sub

Private Sub GenerateThumb()

 Dim i&
 ReDim lImg(nCount - 1)
    
'Load image data getting Original Size of Image
 For i = 0 To nCount - 1
  Call GdipLoadImageFromFile(StrConv(mItem(i).strFile, vbUnicode), lImg(i))
  With mItem(i)
   Call GdipGetImageWidth(lImg(i), .orgWidth)
   Call GdipGetImageHeight(lImg(i), .orgHeight)
   CalculateImageRect .cRECT
  End With
 Next
    
 For i = 0 To nCount - 1
  Call GdipGetImageThumbnail(lImg(i), mItem(i).cRECT.Width, mItem(i).cRECT.Height, mItem(i).hImage)
  DoEvents
 Next
    
 For i = 0 To nCount - 1
  Call GdipDisposeImage(lImg(i))
 Next
    
End Sub

Private Sub CalculateImageRect(mRect As IMGRECT)
 On Error Resume Next
 mRect.Width = picThumb(0).Width
 mRect.Height = picThumb(0).Height
 'mRect.Width = tWidth - 6
 'mRect.Height = tHeight - 20
End Sub

Private Sub Draw()

 On Error Resume Next
 Dim i&, mGraphic&
 Dim mRect As RECT
    
 For i = 0 To nCount - 1
         
  picThumb(i).Cls
  Call GdipCreateFromHDC(picThumb(i).hDC, mGraphic)
  
  With mItem(i)
   Call GdipDrawImageRect(mGraphic, .hImage, 0, 0, .cRECT.Width, .cRECT.Height)
   picText(i).CurrentX = (picText(i).Width - picText(i).TextWidth(flName(i))) / 2
   picText(i).CurrentY = 2
   picText(i).ForeColor = m_ForeColorText
   picText(i).Print flName(i)
  End With
  picThumb(i).Refresh
  picText(i).Refresh
  Call GdipDeleteGraphics(mGraphic)
 Next
 
End Sub

Private Sub FillRectEx(ByVal hDC&, X&, Y&, Width&, Height&, lColor&)

 Dim hBrush&, hResult&, mRect As RECT
 hBrush = CreateSolidBrush(lColor)
 SetRect mRect, X, Y, X + Width, Y + Height
 hResult = FillRect(hDC, mRect, hBrush)
'Clean up
 DeleteObject hBrush
 DeleteObject hResult
 
End Sub

Private Function GetCellNo(tIndex As Integer) As Integer
 
 Dim totThumbWidth As Long
 Dim totThumbGap As Long
 Dim oneThumb As Long
 
 totThumbWidth = tWidth * m_Cols
 totThumbGap = m_Cols * 8
 oneThumb = (totThumbWidth + totThumbGap) / m_Cols
 GetCellNo = (picBack(tIndex).Left + tWidth) / oneThumb
 
End Function

Private Sub ChangeTextProp()
 
 On Error Resume Next
 Dim i&
 Dim mRect As RECT
     
 For i = 0 To nCount - 1
  With mItem(i)
   picText(i).BackColor = m_BackColorText
   picText(i).CurrentX = (picText(i).Width - picText(i).TextWidth(flName(i))) / 2
   picText(i).CurrentY = 2
   picText(i).ForeColor = m_ForeColorText
   picText(i).Print flName(i)
  End With
  picText(i).Refresh
 Next
 
End Sub

Private Sub SetFrame(thSize As ThumbnailSizeEnum)
 If thSize = Tinny Then
  tHeight = 86
  tWidth = 96
 ElseIf thSize = Small Then
  tHeight = 92
  tWidth = 112
 ElseIf thSize = Medium Then
  tHeight = 98
  tWidth = 128
 ElseIf thSize = Large Then
  tHeight = 104
  tWidth = 144
 ElseIf thSize = [Extra Large] Then
  tHeight = 110
  tWidth = 160
 End If
End Sub

Private Sub RemoveThumbs()
 Dim lIdx As Long
 For lIdx = 1 To nCount - 1
  Unload picThumb(lIdx)
  Unload picText(lIdx)
  Unload picBack(lIdx)
 Next lIdx
End Sub

Private Function OnArea() As Boolean

 Dim mpos As POINTAPI
 Dim oRect As RECT
 
 GetCursorPos mpos
 GetWindowRect Me.hwnd, oRect
 
 If mpos.X >= oRect.Left And mpos.X <= oRect.Right And _
  mpos.Y >= oRect.Top And mpos.Y <= oRect.Bottom Then
  OnArea = True
 Else
  OnArea = False
 End If
 
End Function

Public Sub ScrollDown()
Attribute ScrollDown.VB_MemberFlags = "40"

 If OnArea = True Then
  If VScroll.Value <= VScroll.Max - VScroll.SmallChange Then
   VScroll.Value = VScroll.Value + VScroll.SmallChange
  ElseIf VScroll.Value >= VScroll.Max Then
   VScroll.Value = VScroll.Max
  End If
 End If
 
End Sub

Public Sub ScrollUp()
Attribute ScrollUp.VB_MemberFlags = "40"

 If OnArea = True Then
  If VScroll.Value >= VScroll.SmallChange Then
   VScroll.Value = VScroll.Value - VScroll.SmallChange
  ElseIf VScroll.Value = VScroll.Min Then
   VScroll.Value = VScroll.Min
  End If
 End If

End Sub

Private Sub ProcessMessages()
 
 On Error GoTo Err_Site
 Dim Message As msg
 Dim ctrl As Control
 
 Do While Not bCancel
 
  If Not UserControl.Ambient.UserMode = True Then Exit Do 'If In Deign Mode
  WaitMessage  'Wait For message
  'if the mousewheel is used:
  If PeekMessage(Message, UserControl.Parent.hwnd, WM_MOUSEWHEEL, WM_MOUSEWHEEL, PM_REMOVE) Then
   If Message.wParam < 0 Then 'scroll up
    For Each ctrl In UserControl.Parent.Controls
     If TypeOf ctrl Is dsThumb Then
      ctrl.ScrollDown
     End If
    Next
   Else        'scroll down
    For Each ctrl In UserControl.Parent.Controls
     If TypeOf ctrl Is dsThumb Then
      ctrl.ScrollUp
     End If
    Next
   End If
  End If
  DoEvents
   
 Loop
 
Err_Site:
 If Err.Number = 398 Then bCancel = True
    
End Sub
