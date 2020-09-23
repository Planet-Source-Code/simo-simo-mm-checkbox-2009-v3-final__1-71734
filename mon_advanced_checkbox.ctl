VERSION 5.00
Begin VB.UserControl mm_checkbox 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   1110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   ClipBehavior    =   0  'None
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   ScaleHeight     =   74
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   137
   ToolboxBitmap   =   "mon_advanced_checkbox.ctx":0000
   Begin VB.PictureBox pic_des_small_uncheck_avec_caption 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1530
      Picture         =   "mon_advanced_checkbox.ctx":0312
      ScaleHeight     =   225
      ScaleWidth      =   375
      TabIndex        =   5
      Top             =   1815
      Width           =   375
   End
   Begin VB.PictureBox pic_des_small_check_avec_caption 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   1035
      Picture         =   "mon_advanced_checkbox.ctx":07C8
      ScaleHeight     =   225
      ScaleWidth      =   375
      TabIndex        =   4
      Top             =   1800
      Width           =   375
   End
   Begin VB.PictureBox picarcsmall 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   630
      Picture         =   "mon_advanced_checkbox.ctx":0C7E
      ScaleHeight     =   225
      ScaleWidth      =   45
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   45
   End
   Begin VB.PictureBox picarc 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   435
      Left            =   1530
      Picture         =   "mon_advanced_checkbox.ctx":0D74
      ScaleHeight     =   435
      ScaleWidth      =   105
      TabIndex        =   2
      Top             =   180
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.PictureBox picbig 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   780
      Picture         =   "mon_advanced_checkbox.ctx":106E
      ScaleHeight     =   435
      ScaleWidth      =   720
      TabIndex        =   1
      Top             =   180
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.PictureBox picsmall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   225
      Picture         =   "mon_advanced_checkbox.ctx":2100
      ScaleHeight     =   225
      ScaleWidth      =   345
      TabIndex        =   0
      Top             =   750
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.Image pic_des_big_check_avec_caption 
      Height          =   435
      Left            =   120
      Picture         =   "mon_advanced_checkbox.ctx":257A
      Top             =   2730
      Width           =   780
   End
   Begin VB.Image pic_des_big_uncheck_avec_caption 
      Height          =   435
      Left            =   945
      Picture         =   "mon_advanced_checkbox.ctx":3768
      Top             =   2745
      Width           =   780
   End
   Begin VB.Image pic_des_small_check 
      Height          =   225
      Left            =   210
      Picture         =   "mon_advanced_checkbox.ctx":4956
      Top             =   1785
      Width           =   345
   End
   Begin VB.Image pic_des_small_uncheck 
      Height          =   225
      Left            =   600
      Picture         =   "mon_advanced_checkbox.ctx":4DD0
      Top             =   1800
      Width           =   345
   End
   Begin VB.Image pic_des_big_uncheck 
      Height          =   435
      Left            =   915
      Picture         =   "mon_advanced_checkbox.ctx":524A
      Top             =   2130
      Width           =   720
   End
   Begin VB.Image pic_des_big_check 
      Height          =   435
      Left            =   135
      Picture         =   "mon_advanced_checkbox.ctx":62DC
      Top             =   2115
      Width           =   720
   End
End
Attribute VB_Name = "mm_checkbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
Option Explicit

'EVENTS.
Public Event Click()
Public Event DoubleClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnters(ByVal X As Long, ByVal Y As Long)
Public Event MouseLeaves(ByVal X As Long, ByVal Y As Long)


Private udtPoint As POINTAPI
Private bolMouseDown As Boolean
Private bolMouseOver As Boolean
'Private bolHasFocus As Boolean
Private bolEnabled As Boolean
Private bolChecked As Boolean
Private bolSmall As Boolean
Private lonRoundValue As Long
Private lonRect As Long
Private button_clique As Integer

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Dim mon_rect As RECT

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long


'pour le gradient (le petit circle)
Dim AA1 As New LineGS 'DrawRadial

Private m_Activecolor As OLE_COLOR
Private m_desActivecolor As OLE_COLOR
Private m_Caption As String
Private fntFont As Font 'Caption font.
Private m_CaptionColor As OLE_COLOR


Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Const DT_WORDBREAK = &H10
Private Const DT_CENTER = &H1 Or DT_WORDBREAK Or &H4

Sub mon_gradient(mcolor As Long, X As Integer, Y As Integer, iCircle As Integer)
   Dim I As Integer
       
   'UserControl.Cls
   UserControl.DrawStyle = 5
   UserControl.FillStyle = 0
                                                      '|                                            |
    
      With UserControl
         'Copier DIBits dans un array
         AA1.DIB .hdc, .Image.Handle, .ScaleWidth, .ScaleHeight
      End With
   
   '1er cercle en gris
    If Not Small Then
        'bordure
        'For I = 1 To 2
        '    AA1.CircleDIB UserControl.ScaleWidth / 2, UserControl.ScaleHeight / 2, UserControl.ScaleWidth - 28 - I, UserControl.ScaleHeight - 15 - I, vbWhite '&HDAD4CE
        'Next I
        For I = 5 To 6
            'AA1.DIB .hdc, .Image.Handle, .ScaleWidth, .ScaleHeight
            AA1.CircleDIB X, Y, iCircle + I, iCircle + I, &HDAD4CE  'RGB(128, 128, 128) 'vbRed ''100, 100, I * 0.75, I * 0.75, vbRed
            'AA1.Array2Pic
        Next I
    Else
        'For I = 3 To 4
            AA1.CircleDIB X, Y, iCircle + 3, iCircle + 3, &HDAD4CE  'RGB(128, 128, 128) 'vbRed ''100, 100, I * 0.75, I * 0.75, vbRed
            'AA1.Array2Pic
        'Next I
    End If
   
'    'simulate a circle with blendcolor
    AA1.CircleDIB X, Y, iCircle + 1, iCircle + 1, BlendColor(mcolor, vbWhite, 100) '&HDAD4CE    'RGB(128, 128, 128) 'vbRed ''100, 100, I * 0.75, I * 0.75, vbRed
    AA1.CircleDIB X, Y, iCircle + 2, iCircle + 2, BlendColor(mcolor, vbWhite, 50) '&HDAD4CE    'RGB(128, 128, 128) 'vbRed ''100, 100, I * 0.75, I * 0.75, vbRed

      For I = iCircle To 0 Step -1
        AA1.CircleDIB X, Y, I, I, BlendColor(mcolor, vbWhite, I * (255 / iCircle))
     Next I
     
     'refresh picture for usercontrol
      AA1.Array2Pic
      
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
    dlgAbout.Show 1
End Sub


Private Function PointInControl(X As Single, Y As Single) As Boolean
  If X >= 0 And X <= UserControl.ScaleWidth And _
    Y >= 0 And Y <= UserControl.ScaleHeight Then
    PointInControl = True
  End If
End Function

Private Sub PaintControl()
    
Dim rc As RECT

    UserControl.Refresh
    UserControl.Picture = LoadPicture("")
    UserControl.Refresh
    UserControl.Cls
    
    'If bolEnabled Then
        pic_des_small_check.Top = -200
        pic_des_small_uncheck.Top = -200
        pic_des_big_check.Top = -200
        pic_des_big_uncheck.Top = -200
        pic_des_small_check_avec_caption.Top = -200
        pic_des_small_uncheck_avec_caption.Top = -200
        pic_des_big_uncheck_avec_caption.Top = -200
        pic_des_big_check_avec_caption.Top = -200
        
    'Else
    If Not bolEnabled Then
        If bolSmall Then
            If Checked Then
                If m_Caption <> "" Then
                    pic_des_small_check_avec_caption.Top = 0
                    pic_des_small_check_avec_caption.Left = 0
                Else
                    pic_des_small_check.Top = 0
                    pic_des_small_check.Left = 0
                End If
            Else
                If m_Caption <> "" Then
                    pic_des_small_uncheck_avec_caption.Top = 0
                    pic_des_small_uncheck_avec_caption.Left = 0
                Else
                    pic_des_small_uncheck.Top = 0
                    pic_des_small_uncheck.Left = 0
                End If
            End If
        Else
            If Checked Then
                If m_Caption <> "" Then
                    pic_des_big_check_avec_caption.Top = 0
                    pic_des_big_check_avec_caption.Left = 0
                Else
                    pic_des_big_check.Top = 0
                    pic_des_big_check.Left = 0
                End If
            Else
                If m_Caption <> "" Then
                    pic_des_big_uncheck_avec_caption.Top = 0
                    pic_des_big_uncheck_avec_caption.Left = 0
                Else
                    pic_des_big_uncheck.Top = 0
                    pic_des_big_uncheck.Left = 0
                End If
            End If
        End If
    End If
    '-----------------------------------------
    
    If Small Then
        UserControl_Resize
        UserControl.Cls
        If m_Caption <> "" Then
            'Center stretch
            StretchBlt UserControl.hdc, 0, 0, ScaleWidth, ScaleHeight, picsmall.hdc, 5, 0, 1, ScaleHeight, vbSrcCopy
            'Left
            'StretchBlt UserControl.hdc, 0, 0, 10, ScaleHeight, picsmall.hdc, 0, 0, 10, ScaleHeight, vbSrcCopy
            BitBlt UserControl.hdc, 0, 0, 6, ScaleHeight, picsmall.hdc, 0, 0, vbSrcCopy
            'Right
            BitBlt UserControl.hdc, ScaleWidth - 10, 0, 9, ScaleHeight, picsmall.hdc, picsmall.Width - 9, 0, vbSrcCopy
            'end of checkbox
            If bolEnabled Then BitBlt UserControl.hdc, picsmall.Width - 1, 0, picarcsmall.Width, ScaleHeight, picarcsmall.hdc, 0, 0, vbSrcCopy
            'draw caption
            rc.Left = picsmall.Width + 6: rc.Top = (picsmall.Height - (picsmall.TextHeight("-") / Screen.TwipsPerPixelY)) / 2 '8
            rc.Right = UserControl.ScaleWidth: rc.Bottom = UserControl.ScaleHeight
            If bolEnabled Then
                UserControl.ForeColor = m_CaptionColor 'm_Activecolor
            Else
                UserControl.ForeColor = vbGrayText
            End If
            DrawText UserControl.hdc, m_Caption, Len(m_Caption), rc, 0 ', DT_CENTER
        Else
            UserControl.Picture = picsmall.Picture
        End If
        
        UserControl.Refresh
        
        If Checked Then
            mon_gradient m_Activecolor, 7, (UserControl.ScaleHeight / 2) - 1, 3
        Else
            mon_gradient m_desActivecolor, picsmall.Width - 8, (UserControl.ScaleHeight / 2) - 1, 3 '6
        End If
    Else
        UserControl_Resize
        UserControl.Cls
        If m_Caption <> "" Then
            'center stretch
            StretchBlt UserControl.hdc, 0, 0, ScaleWidth, ScaleHeight, picbig.hdc, 20, 0, 2, ScaleHeight, vbSrcCopy
            'Left
            BitBlt UserControl.hdc, 0, 0, 15, ScaleHeight, picbig.hdc, 0, 0, vbSrcCopy
            'Right
            BitBlt UserControl.hdc, ScaleWidth - 16, 0, 15, ScaleHeight, picbig.hdc, picbig.Width - 15, 0, vbSrcCopy
            'End of checkbox
            If bolEnabled Then BitBlt UserControl.hdc, picbig.Width - 3, 0, picarc.Width, ScaleHeight, picarc.hdc, 0, 0, vbSrcCopy
            'draw caption
            rc.Left = picbig.Width + 8: rc.Top = (picbig.Height - (picbig.TextHeight("-") / Screen.TwipsPerPixelY)) / 2 '8 'picbig.Height / 2
            rc.Right = UserControl.ScaleWidth: rc.Bottom = UserControl.ScaleHeight
            If bolEnabled Then
                UserControl.ForeColor = m_CaptionColor 'm_Activecolor
            Else
                UserControl.ForeColor = vbGrayText 'vbButtonShadow
            End If
            DrawText UserControl.hdc, m_Caption, Len(m_Caption), rc, 0 ', DT_CENTER
        Else
            UserControl.Picture = picbig.Picture
        End If
        
        UserControl.Refresh
        
        If Checked Then
            mon_gradient m_Activecolor, 15, (UserControl.ScaleHeight / 2) - 1, 8
        Else
            mon_gradient m_desActivecolor, picbig.Width - 16, (picbig.Height / 2) - 1, 8 '9
        End If
    
    End If
    
End Sub

Public Property Get Activecolor() As OLE_COLOR
   Activecolor = m_Activecolor
End Property
Public Property Get CaptionColor() As OLE_COLOR
   CaptionColor = m_CaptionColor
End Property
Public Property Get desActivecolor() As OLE_COLOR
   desActivecolor = m_desActivecolor
End Property
Public Property Let Activecolor(ByVal New_Activecolor As OLE_COLOR)
   m_Activecolor = New_Activecolor
   PropertyChanged "Activecolor"
   PaintControl
End Property
Public Property Let CaptionColor(ByVal New_CaptionColor As OLE_COLOR)
   m_CaptionColor = New_CaptionColor
   PropertyChanged "CaptionColor"
   PaintControl
End Property
Public Property Get Font() As Font
Set Font = fntFont
End Property
Public Property Set Font(ByVal NewValue As Font)
Set fntFont = NewValue
Set UserControl.Font = NewValue

Set picbig.Font = UserControl.Font
Set picsmall.Font = UserControl.Font


PropertyChanged "Font"
PaintControl
End Property
Public Property Let desActivecolor(ByVal New_desActivecolor As OLE_COLOR)
   m_desActivecolor = New_desActivecolor
   PropertyChanged "desActivecolor"
   PaintControl
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Button Enabled/Disable."
Enabled = bolEnabled
End Property


Public Property Get Small() As Boolean
Small = bolSmall
End Property
Public Property Get Checked() As Boolean
Checked = bolChecked
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
bolEnabled = NewValue
PropertyChanged "Enabled"

UserControl.Enabled = bolEnabled

PaintControl
End Property


Public Property Let Small(ByVal NewValue As Boolean)
bolSmall = NewValue
PropertyChanged "Small"

PaintControl


If Small = True Then
    RoundedValue = 9 '10
Else
    RoundedValue = 26
End If


End Property
Public Property Let Checked(ByVal NewValue As Boolean)
bolChecked = NewValue
PropertyChanged "Checked"

PaintControl


End Property
Public Property Get RoundedValue() As Long
Attribute RoundedValue.VB_Description = "Button Border Rounded Value."
RoundedValue = lonRoundValue
End Property

Public Property Let RoundedValue(ByVal NewValue As Long)


lonRoundValue = NewValue
PropertyChanged "RoundedValue"


UserControl_Resize

End Property

Private Sub UserControl_Click()
If bolEnabled = True Then
    If button_clique = 1 Then
        
        Checked = Not Checked
        'PaintControl
        
        RaiseEvent Click
        RaiseEvent MouseLeaves(0, 0)
    End If
End If
End Sub

Private Sub UserControl_Initialize()
m_Activecolor = &H8000&
m_desActivecolor = &H808080
m_Caption = ""
m_CaptionColor = vbBlack
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bolEnabled = True Then
    button_clique = Button
    If Button = 1 Then
        bolMouseDown = True
        RaiseEvent MouseDown(Button, Shift, X, Y)
'        PaintControl
    End If
End If

End Sub
Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If bolEnabled = False Then Exit Sub
    RaiseEvent MouseMove(Button, Shift, X, Y)
    SetCapture hWnd
    If PointInControl(X, Y) Then
        'pointer on control
        If Not bolMouseOver Then
            bolMouseOver = True
            RaiseEvent MouseEnters(udtPoint.X, udtPoint.Y)
        End If
    Else
        'pointer out of control
        bolMouseOver = False
        bolMouseDown = False
        ReleaseCapture
        RaiseEvent MouseLeaves(udtPoint.X, udtPoint.Y)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If bolEnabled = True Then
    button_clique = Button
    If Button = 1 Then
        RaiseEvent MouseUp(Button, Shift, X, Y)
        bolMouseDown = False
    End If
End If
End Sub

Private Sub UserControl_Paint()
'PaintControl
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
'On Error Resume Next
With PropBag
    
    Let Enabled = .ReadProperty("Enabled", True)
    Let Checked = .ReadProperty("Checked", False)
    Let Small = .ReadProperty("Small", True)
    Let RoundedValue = .ReadProperty("RoundedValue", 5)
    Let Activecolor = .ReadProperty("Activecolor", m_Activecolor) ' &H117B28) 'vbGreen)
    Let desActivecolor = .ReadProperty("desActivecolor", m_desActivecolor) ' &H117B28) 'vbGreen)
    
    Let Caption = .ReadProperty("Caption", "")
    Set Font = .ReadProperty("Font", Ambient.Font)
    Let CaptionColor = .ReadProperty("CaptionColor", m_CaptionColor)
End With
End Sub
Private Sub UserControl_Resize()
    
    If Small Then
        'UserControl.Width = (picsmall.Width + 1) * Screen.TwipsPerPixelX
        If m_Caption <> "" Then
            UserControl.Width = ((picsmall.Width + 1) * Screen.TwipsPerPixelX) + (picsmall.TextWidth(m_Caption) + 300)
        Else
            UserControl.Width = (picsmall.Width + 1) * Screen.TwipsPerPixelX '* 2
        End If
        UserControl.Height = (picsmall.Height + 1) * Screen.TwipsPerPixelY
    Else
        If m_Caption <> "" Then
            UserControl.Width = ((picbig.Width + 1) * Screen.TwipsPerPixelX) + (picbig.TextWidth(m_Caption) + 300)
        Else
            UserControl.Width = (picbig.Width + 1) * Screen.TwipsPerPixelX '* 2
        End If
        UserControl.Height = (picbig.Height + 1) * Screen.TwipsPerPixelY
    End If
    

lonRect = CreateRoundRectRgn(0, 0, ScaleWidth, ScaleHeight, lonRoundValue, lonRoundValue)     '- 1
SetWindowRgn UserControl.hWnd, lonRect, True



End Sub

Private Sub UserControl_Terminate()
bolMouseDown = False
bolMouseOver = False
'bolHasFocus = False
'UserControl.Cls
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
'On Error Resume Next
With PropBag
    .WriteProperty "Enabled", bolEnabled, True
    .WriteProperty "Checked", bolChecked, False
    .WriteProperty "Small", bolSmall, True
    .WriteProperty "RoundedValue", lonRoundValue, 5
    .WriteProperty "Activecolor", m_Activecolor, &H8000& '&H117B28 'vbGreen
    .WriteProperty "desActivecolor", m_desActivecolor, &H808080 '&H94A392
    
    .WriteProperty "Caption", m_Caption, ""
    .WriteProperty "Font", fntFont, Ambient.Font
    .WriteProperty "CaptionColor", m_CaptionColor, vbBlack
End With
End Sub
Private Sub UserControl_InitProperties()
Let Enabled = True
Let Checked = False
Let Small = False 'True
Let RoundedValue = 27 '26 '5

m_Activecolor = &H8000&
m_desActivecolor = &H808080
m_CaptionColor = vbBlack
Set Font = UserControl.Font '"tahoma" 'Ambient.Font

End Sub

Public Property Get Caption() As String
Caption = m_Caption
End Property

Public Property Let Caption(ByVal New_Caption As String)
   m_Caption = New_Caption
   PropertyChanged "Caption"
   PaintControl
End Property
