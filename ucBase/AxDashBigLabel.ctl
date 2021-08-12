VERSION 5.00
Begin VB.UserControl AxDBigLabel 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H008D4214&
   ClientHeight    =   1125
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4230
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForwardFocus    =   -1  'True
   KeyPreview      =   -1  'True
   ScaleHeight     =   75
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   282
   ToolboxBitmap   =   "AxDashBigLabel.ctx":0000
   Begin VB.Timer tmrEffect 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   75
      Top             =   75
   End
End
Attribute VB_Name = "AxDBigLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'-UC-VB6-----------------------------
'UC Name  : AxGButtonLabel
'Version  : 0.01
'Editor   : David Rojas [AxioUK]
'Date     : 27/06/2021
'------------------------------------
Option Explicit

Private Declare Function MulDiv Lib "kernel32.dll" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Declare Function TlsGetValue Lib "kernel32.dll" (ByVal dwTlsIndex As Long) As Long
'Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
'-
'Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long

'Private Declare Function ReleaseCapture Lib "User32" () As Long
'Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long

'Private Declare Function LoadCursor Lib "User32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
'Private Declare Function DestroyCursor Lib "User32" (ByVal hCursor As Long) As Long
'Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As Any, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long

Private Declare Function GdiplusStartup Lib "GdiPlus.dll" (Token As Long, inputbuf As GDIPlusStartupInput, Optional ByVal outputbuf As Long = 0) As Long
Private Declare Sub GdiplusShutdown Lib "GdiPlus.dll" (ByVal Token As Long)
Private Declare Function GdipCreateLineBrushFromRectWithAngleI Lib "GdiPlus.dll" (ByRef mRect As RECTL, ByVal mColor1 As Long, ByVal mColor2 As Long, ByVal mAngle As Single, ByVal mIsAngleScalable As Long, ByVal mWrapMode As Long, ByRef mLineGradient As Long) As Long
'Private Declare Function GdipDrawRectangleI Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal pen As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GdipCreateFromHDC Lib "GdiPlus.dll" (ByVal mhDC As Long, ByRef mGraphics As Long) As Long
Private Declare Function GdipCreatePen1 Lib "GdiPlus.dll" (ByVal mColor As Long, ByVal mWidth As Single, ByVal mUnit As Long, ByRef mPen As Long) As Long
Private Declare Function GdipDeleteGraphics Lib "GdiPlus.dll" (ByVal mGraphics As Long) As Long
Private Declare Function GdipDeleteBrush Lib "GdiPlus.dll" (ByVal brush As Long) As Long
Private Declare Function GdipDeletePen Lib "GdiPlus.dll" (ByVal mPen As Long) As Long
Private Declare Function GdipCreatePath Lib "GdiPlus.dll" (ByRef mBrushMode As Long, ByRef mPath As Long) As Long
'Private Declare Function GdipAddPathLineI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
Private Declare Function GdipAddPathArcI Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mX As Long, ByVal mY As Long, ByVal mWidth As Long, ByVal mHeight As Long, ByVal mStartAngle As Single, ByVal mSweepAngle As Single) As Long
Private Declare Function GdipClosePathFigures Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDeletePath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
Private Declare Function GdipDrawPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mPath As Long) As Long
Private Declare Function GdipFillPath Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mBrush As Long, ByVal mPath As Long) As Long
Private Declare Function GdipSetSmoothingMode Lib "GdiPlus.dll" (ByVal graphics As Long, ByVal SmoothingMd As Long) As Long
Private Declare Function GdipCreateSolidFill Lib "gdiplus" (ByVal ARGB As Long, ByRef brush As Long) As Long
Private Declare Function GdipAddPathString Lib "GdiPlus.dll" (ByVal mPath As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFamily As Long, ByVal mStyle As Long, ByVal mEmSize As Single, ByRef mLayoutRect As RECTS, ByVal mFormat As Long) As Long
'Private Declare Function GdipMeasureString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByRef mBoundingBox As RECTS, ByRef mCodepointsFitted As Long, ByRef mLinesFilled As Long) As Long
'Private Declare Function GdipCreateFont Lib "GdiPlus.dll" (ByVal mFontFamily As Long, ByVal mEmSize As Single, ByVal mStyle As Long, ByVal mUnit As Long, ByRef mFont As Long) As Long
'Private Declare Function GdipDeleteFont Lib "GdiPlus.dll" (ByVal mFont As Long) As Long
Private Declare Function GdipCreateFontFamilyFromName Lib "gdiplus" (ByVal Name As Long, ByVal fontCollection As Long, fontFamily As Long) As Long
Private Declare Function GdipDeleteFontFamily Lib "gdiplus" (ByVal fontFamily As Long) As Long
Private Declare Function GdipGetGenericFontFamilySansSerif Lib "GdiPlus.dll" (ByRef mNativeFamily As Long) As Long
'Private Declare Function GdipDrawString Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mString As Long, ByVal mLength As Long, ByVal mFont As Long, ByRef mLayoutRect As RECTS, ByVal mStringFormat As Long, ByVal mBrush As Long) As Long
Private Declare Function GdipSetStringFormatTrimming Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mTrimming As eStringTrimming) As Long
Private Declare Function GdipCreateStringFormat Lib "gdiplus" (ByVal formatAttributes As Long, ByVal language As Integer, StringFormat As Long) As Long
Private Declare Function GdipSetStringFormatFlags Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mFlags As eStringFormatFlags) As Long
Private Declare Function GdipSetStringFormatAlign Lib "gdiplus" (ByVal StringFormat As Long, ByVal Align As eStringAlignment) As Long
Private Declare Function GdipSetStringFormatLineAlign Lib "GdiPlus.dll" (ByVal mFormat As Long, ByVal mAlign As eStringAlignment) As Long
Private Declare Function GdipDeleteStringFormat Lib "GdiPlus.dll" (ByVal mFormat As Long) As Long
'Private Declare Function GdipDrawLineI Lib "GdiPlus.dll" (ByVal mGraphics As Long, ByVal mPen As Long, ByVal mX1 As Long, ByVal mY1 As Long, ByVal mX2 As Long, ByVal mY2 As Long) As Long
'Private Declare Function GdipDrawPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal pen As Long, ByRef pPoints As Any, ByVal count As Long) As Long
'Private Declare Function GdipFillPolygonI Lib "gdiplus" (ByVal graphics As Long, ByVal brush As Long, ByRef pPoints As Any, ByVal count As Long, ByVal FillMode As Long) As Long
Private Declare Function GdipTranslateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal dX As Single, ByVal dY As Single, ByVal Order As Long) As Long
Private Declare Function GdipRotateWorldTransform Lib "gdiplus" (ByVal graphics As Long, ByVal Angle As Single, ByVal Order As Long) As Long
Private Declare Function GdipResetWorldTransform Lib "GdiPlus.dll" (ByVal graphics As Long) As Long
'Private Declare Function GdipResetPath Lib "GdiPlus.dll" (ByVal mPath As Long) As Long
'---
'---
'Private Declare Function DrawTextW Lib "user32.dll" (ByVal hdc As Long, lpStr As Long, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTL) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
'Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECTS, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

'Private Type RECT
'  Left As Long
'  Top As Long
'  Right As Long
'  Bottom As Long
'End Type

'Private Type POINTAPI
'    x As Long
'    y As Long
'End Type

'Private Type POINTS
'   x As Single
'   y As Single
'End Type

Private Enum GDIPLUS_FONTSTYLE
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

'Private Const LF_FACESIZE = 32
'Private Const SYSTEM_FONT = 13
'Private Const OBJ_FONT As Long = 6&

Private Type GDIPlusStartupInput
    GdiPlusVersion           As Long
    DebugEventCallback       As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs   As Long
End Type

'Private Enum HotkeyPrefix
'    HotkeyPrefixNone = &H0
'    HotkeyPrefixShow = &H1
'    HotkeyPrefixHide = &H2
'End Enum

Private Type POINTL
    x As Long
    y As Long
End Type

Private Type RECTL
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type RECTS
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

'Private Type PicBmp
'  Size As Long
'  type As Long
'  hBmp As Long
'  hpal As Long
'  Reserved As Long
'End Type

'Private Enum PenAlignment
'    PenAlignmentCenter = &H0
'    PenAlignmentInset = &H1
'End Enum

'EVENTS------------------------------------
Public Event Click()
'Public Event ChangeValue(ByVal Value As Boolean)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
''-----------------------------------------

'Private Const CombineModeExclude As Long = &H4
Private Const WrapModeTileFlipXY = &H3
'Private Const SmoothingModeHighQuality As Long = &H2
Private Const SmoothingModeAntiAlias As Long = &H4
Private Const LOGPIXELSX As Long = 88
Private Const LOGPIXELSY As Long = 90
Private Const TLS_MINIMUM_AVAILABLE As Long = 64
'Private Const CLR_INVALID = -1
'Private Const WM_NCLBUTTONDOWN = &HA1
'Private Const HTCAPTION = 2
'Private Const IDC_HAND As Long = 32649
Private Const UnitPixel As Long = &H2&

'Default Property Values:
Const m_def_ForeColor = 0
Const m_def_Appearance = 0
Const m_def_BackStyle = 0
Const m_def_BorderStyle = 0
Const m_def_Color1 = &HD59B5B
Const m_def_Color2 = &H6A5444
Const m_def_Angulo = 0

'Property Variables:
Dim GdipToken As Long
Dim nScale    As Single
'Dim hCur      As Long
Dim hFontCollection As Long
Dim hGraphics As Long
'Dim hPen      As Long
'Dim hBrush    As Long

Dim m_BorderColor   As OLE_COLOR
Dim m_ForeColor     As OLE_COLOR
Dim m_ForeColor2    As OLE_COLOR
Dim m_ForeColor3    As OLE_COLOR
Dim m_Color1        As OLE_COLOR
Dim m_Color2        As OLE_COLOR
Dim m_BorderWidth   As Long
Dim m_Enabled       As Boolean
Dim m_Angulo        As Single
Dim m_CornerCurve   As Long
Dim m_Caption3Enable As Boolean
Dim m_Caption       As String
Dim m_Caption2      As String
Dim m_Caption3      As String
'Dim m_Top         As Long
'Dim m_Left        As Long
Dim m_Opacity     As Long
Dim cl_hWnd       As Long

Dim m_OldBorderColor As OLE_COLOR
Dim m_BoxedColor As OLE_COLOR
Dim m_ChangeBorderOnFocus As Boolean
'Dim m_OnFocus As Boolean

Private m_Font          As StdFont
Private m_Font2         As StdFont
Private m_Font3         As StdFont
Private m_IconFont      As StdFont
Private m_IconCharCode  As Long
Private m_IconForeColor As Long
Private m_IconAlignH    As eTextAlignH
Private m_IconAlignV    As eTextAlignV

Private m_CaptionAlignV As eTextAlignV
Private m_CaptionAlignH As eTextAlignH
Private m_Caption2AlignV As eTextAlignV
Private m_Caption2AlignH As eTextAlignH
Private m_Caption3AlignV As eTextAlignV
Private m_Caption3AlignH As eTextAlignH

Private m_EffectFade  As Boolean
Private m_InitialOpacity As Long
Private m_Transparent As Boolean

Dim m_IconBox       As eBoxed

Dim m_isBoxed      As Boolean
Dim m_Clicked      As Boolean
Dim m_Clickable    As Boolean
Dim m_Filled       As Boolean


Public Sub CopyAmbient()
Dim OPic As StdPicture

On Error GoTo Err

    With UserControl
        Set .Picture = Nothing
        Set OPic = Extender.Container.Image
        .BackColor = Extender.Container.BackColor
        UserControl.PaintPicture OPic, 0, 0, , , Extender.Left, Extender.Top
        Set .Picture = .Image
    End With
Err:
End Sub

Public Sub Refresh()
  UserControl.Cls
  CopyAmbient
  Draw
End Sub

'Private Function AddIconChar(Rct As RECT, IconColor As OLE_COLOR)
''Dim Rct As RECT
'Dim pFont       As IFont
'Dim lFontOld    As Long
'
'On Error GoTo ErrF
'With UserControl
'  .AutoRedraw = True
'  Set pFont = IconFont
'  lFontOld = SelectObject(.hdc, pFont.hFont)
'
'  .ForeColor = IconColor
'
''  Rct.Left = (IconFont.Size / 2) + m_PadX
''  Rct.Top = (IconFont.Size / 2) + m_PadY
''  Rct.Right = UserControl.ScaleWidth
''  Rct.Bottom = UserControl.ScaleHeight
'
'  DrawTextW .hdc, IconCharCode, 1, Rct, 0
'
'  Call SelectObject(.hdc, lFontOld)
'
'ErrF:
'  Set pFont = Nothing
'End With
'End Function

Private Function ARGB(ByVal RGBColor As Long, ByVal Opacity As Long) As Long
  If (RGBColor And &H80000000) Then RGBColor = GetSysColor(RGBColor And &HFF&)
  ARGB = (RGBColor And &HFF00&) Or (RGBColor And &HFF0000) \ &H10000 Or (RGBColor And &HFF) * &H10000
  Opacity = CByte((Abs(Opacity) / 100) * 255)
  If Opacity < 128 Then
      If Opacity < 0& Then Opacity = 0&
      ARGB = ARGB Or Opacity * &H1000000
  Else
      If Opacity > 255& Then Opacity = 255&
      ARGB = ARGB Or (Opacity - 128&) * &H1000000 Or &H80000000
  End If
End Function

Private Function ChrW2(ByVal CharCode As Long) As String
  Const POW10 As Long = 2 ^ 10
  If CharCode <= &HFFFF& Then ChrW2 = ChrW$(CharCode) Else _
                              ChrW2 = ChrW$(&HD800& + (CharCode And &HFFFF&) \ POW10) & _
                                      ChrW$(&HDC00& + (CharCode And (POW10 - 1)))
End Function

Private Sub Draw()
Dim IcoBox As RECTS
Dim REC As RECTL
Dim BOX As RECTL
Dim stREC As RECTS
Dim stREC2 As RECTS
Dim stREC3 As RECTS
Dim lBorder As Long, mBorder As Long

With UserControl
    
  GdipCreateFromHDC .hdc, hGraphics
  GdipSetSmoothingMode hGraphics, SmoothingModeAntiAlias

  lBorder = m_BorderWidth * 2
  mBorder = lBorder / 2
    
  If m_Caption3Enable Then
    stREC.Left = IIf(m_IconBox = ebLeft, (.ScaleHeight - lBorder) + 2, 1) * nScale
    stREC.Top = mBorder * nScale
    stREC.Width = .ScaleWidth - (.ScaleHeight - lBorder) - 3 * nScale
    stREC.Height = (.ScaleHeight - lBorder) / 3 * nScale
    
    stREC2.Left = IIf(m_IconBox = ebLeft, (.ScaleHeight - lBorder) + 2, 1) * nScale
    stREC2.Top = (.ScaleHeight - lBorder) / 3 * nScale
    stREC2.Width = .ScaleWidth - (.ScaleHeight - lBorder) - 3 * nScale
    stREC2.Height = (.ScaleHeight - lBorder) / 3 * nScale
    
    stREC3.Left = IIf(m_IconBox = ebLeft, (.ScaleHeight - lBorder) + 2, 1) * nScale
    stREC3.Top = ((.ScaleHeight - lBorder) / 3) * 2 * nScale
    stREC3.Width = .ScaleWidth - (.ScaleHeight - lBorder) - 3 * nScale
    stREC3.Height = (.ScaleHeight - lBorder) / 3 * nScale
  Else
    stREC.Left = IIf(m_IconBox = ebLeft, (.ScaleHeight - lBorder) + 2, 1) * nScale
    stREC.Top = mBorder * nScale
    stREC.Width = .ScaleWidth - (.ScaleHeight - lBorder) - 3 * nScale
    stREC.Height = (.ScaleHeight - lBorder) / 2 * nScale
    
    stREC2.Left = IIf(m_IconBox = ebLeft, (.ScaleHeight - lBorder) + 2, 1) * nScale
    stREC2.Top = (.ScaleHeight - lBorder) / 2 * nScale
    stREC2.Width = .ScaleWidth - (.ScaleHeight - lBorder) - 3 * nScale
    stREC2.Height = (.ScaleHeight - lBorder) / 2 * nScale
  End If
  
  REC.Left = 1:     REC.Top = 1
  REC.Width = .ScaleWidth - 2
  REC.Height = .ScaleHeight - 2
          
  Select Case m_IconBox
    Case ebLeft
      BOX.Left = 2 * nScale
      BOX.Top = 2 * nScale
      BOX.Width = .ScaleHeight - 3 * nScale
      BOX.Height = .ScaleHeight - 4 * nScale
    Case ebRight
      BOX.Left = .ScaleWidth - .ScaleHeight * nScale
      BOX.Top = 2 * nScale
      BOX.Width = .ScaleHeight - 2 * nScale
      BOX.Height = .ScaleHeight - 4 * nScale
  End Select
      
  IcoBox.Left = BOX.Left: IcoBox.Top = BOX.Top
  IcoBox.Width = BOX.Width: IcoBox.Height = BOX.Height

  SafeRange m_Opacity, 0, 100

  If m_Clickable Then
    GoTo Clicked
  Else
    GoTo NoClicked
  End If

  '-DRAW BUTTON--------------
NoClicked:
  If m_EffectFade Then
    GoTo Effect
  Else
    gRoundRect hGraphics, REC, ARGB(m_Color1, 100), ARGB(m_Color2, 100), m_Angulo, ARGB(m_BorderColor, 100), m_CornerCurve, m_Filled
    DrawCaption hGraphics, m_Caption, m_Font, stREC, m_ForeColor, 100, 0, m_CaptionAlignH, m_CaptionAlignV, False
    DrawCaption hGraphics, m_Caption2, m_Font2, stREC2, m_ForeColor2, 100, 0, m_Caption2AlignH, m_Caption2AlignV, False
    If m_Caption3Enable Then DrawCaption hGraphics, m_Caption3, m_Font3, stREC3, m_ForeColor3, 100, 0, m_Caption3AlignH, m_Caption3AlignV, False
    If m_isBoxed Then gRoundRect hGraphics, BOX, ARGB(m_BoxedColor, 100), ARGB(m_BoxedColor, 100), m_Angulo, ARGB(m_BoxedColor, 100), m_CornerCurve, m_Filled
    GoTo Continuar
  End If
  
Clicked:
  gRoundRect hGraphics, REC, ARGB(m_Color1, 100), ARGB(m_Color2, 100), m_Angulo, ARGB(m_BorderColor, 100), m_CornerCurve, m_Filled
  DrawCaption hGraphics, m_Caption, m_Font, stREC, m_ForeColor, 100, 0, m_CaptionAlignH, m_CaptionAlignV, False
  DrawCaption hGraphics, m_Caption2, m_Font2, stREC2, m_ForeColor2, 100, 0, m_Caption2AlignH, m_Caption2AlignV, False
  If m_Caption3Enable Then DrawCaption hGraphics, m_Caption3, m_Font3, stREC3, m_ForeColor3, 100, 0, m_Caption3AlignH, m_Caption3AlignV, False
  If m_isBoxed Then gRoundRect hGraphics, BOX, ARGB(m_BoxedColor, 100), ARGB(m_BoxedColor, 100), m_Angulo, ARGB(m_BoxedColor, 100), m_CornerCurve, m_Filled
  GoTo Continuar
  
Effect:
  gRoundRect hGraphics, REC, ARGB(m_Color1, m_Opacity), ARGB(m_Color2, m_Opacity), m_Angulo, ARGB(m_BorderColor, m_Opacity), m_CornerCurve, m_Filled
  DrawCaption hGraphics, m_Caption, m_Font, stREC, m_ForeColor, CLng(m_Opacity), 0, m_CaptionAlignH, m_CaptionAlignV, False
  DrawCaption hGraphics, m_Caption2, m_Font2, stREC2, m_ForeColor2, CLng(m_Opacity), 0, m_Caption2AlignH, m_Caption2AlignV, False
  If m_Caption3Enable Then DrawCaption hGraphics, m_Caption3, m_Font3, stREC3, m_ForeColor3, CLng(m_Opacity), 0, m_Caption3AlignH, m_Caption3AlignV, False
  If m_isBoxed Then gRoundRect hGraphics, BOX, ARGB(m_BoxedColor, m_Opacity), ARGB(m_BoxedColor, m_Opacity), m_Angulo, ARGB(m_BoxedColor, m_Opacity), m_CornerCurve, m_Filled

Continuar:
   DrawCaption hGraphics, IconCharCode, IconFont, IcoBox, m_IconForeColor, CLng(100), 0, m_IconAlignH, m_IconAlignV, True

'  '---------------

'  '---------------
  GdipDeleteGraphics hGraphics
  '---------------
  If m_Transparent Then
    .BackStyle = 0
    .MaskColor = .BackColor
    Set .MaskPicture = .Image
  End If
  '---------------
End With

End Sub

Private Function DrawCaption(ByVal hGraphics As Long, sString As Variant, oFont As StdFont, layoutRect As RECTS, _
                             TextColor As OLE_COLOR, ColorOpacity As Integer, mAngle As Single, HAlign As eTextAlignH, _
                             VAlign As eTextAlignV, Icon As Boolean) As Long
Dim hPath As Long
Dim hBrush As Long
Dim hFontFamily As Long
Dim hFormat As Long
Dim lFontSize As Long
Dim lFontStyle As GDIPLUS_FONTSTYLE
Dim newY As Long, newX As Long

On Error Resume Next

    If GdipCreatePath(&H0, hPath) = 0 Then

        If GdipCreateStringFormat(0, 0, hFormat) = 0 Then
            GdipSetStringFormatFlags hFormat, StringFormatFlagsNoWrap 'bWordWrap
            GdipSetStringFormatTrimming hFormat, StringTrimmingEllipsisWord
            GdipSetStringFormatAlign hFormat, HAlign
            GdipSetStringFormatLineAlign hFormat, VAlign
        End If

        GetFontStyleAndSize oFont, lFontStyle, lFontSize

        If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), 0, hFontFamily) Then
            If hFontCollection Then
                If GdipCreateFontFamilyFromName(StrPtr(oFont.Name), hFontCollection, hFontFamily) Then
                    If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
                End If
            Else
                If GdipGetGenericFontFamilySansSerif(hFontFamily) Then Exit Function
            End If
        End If
'------------------------------------------------------------------------
        If mAngle <> 0 Then
            newY = (layoutRect.Height / 2)
            newX = (layoutRect.Width / 2)
            Call GdipTranslateWorldTransform(hGraphics, newX, newY, 0)
            Call GdipRotateWorldTransform(hGraphics, mAngle, 0)
            Call GdipTranslateWorldTransform(hGraphics, -newX, -newY, 0)
        End If
'------------------------------------------------------------------------
      If Icon Then
        GdipAddPathString hPath, StrPtr(ChrW2(sString)), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
      Else
        GdipAddPathString hPath, StrPtr(sString), -1, hFontFamily, lFontStyle, lFontSize, layoutRect, hFormat
      End If
'------------------------------------------------------------------------
        GdipDeleteStringFormat hFormat

        GdipCreateSolidFill ARGB(TextColor, ColorOpacity), hBrush

        GdipFillPath hGraphics, hBrush, hPath
        GdipDeleteBrush hBrush

        If mAngle <> 0 Then GdipResetWorldTransform hGraphics

        GdipDeleteFontFamily hFontFamily
            
        GdipDeletePath hPath
    End If

End Function

Private Function GetFontStyleAndSize(oFont As StdFont, lFontStyle As Long, lFontSize As Long)
On Error GoTo ErrO
    Dim hdc As Long
    lFontStyle = 0
    If oFont.Bold Then lFontStyle = lFontStyle Or FontStyleBold
    If oFont.Italic Then lFontStyle = lFontStyle Or FontStyleItalic
    If oFont.Underline Then lFontStyle = lFontStyle Or FontStyleUnderline
    If oFont.Strikethrough Then lFontStyle = lFontStyle Or FontStyleStrikeout
    
    hdc = GetDC(0&)
    lFontSize = MulDiv(oFont.Size, GetDeviceCaps(hdc, LOGPIXELSY), 72)
    ReleaseDC 0&, hdc
ErrO:
End Function

Private Function GetSafeRound(Angle As Integer, Width As Long, Height As Long) As Integer
    Dim lRet As Integer
    lRet = Angle
    If lRet * 2 > Height Then lRet = Height \ 2
    If lRet * 2 > Width Then lRet = Width \ 2
    GetSafeRound = lRet
End Function

'Private Function GetSystemHandCursor() As Picture
'  Dim Pic As PicBmp
'  Dim IPic As IPicture
'  Dim GUID(0 To 3) As Long
'
'  If hCur Then DestroyCursor hCur: hCur = 0
'
'  hCur = LoadCursor(ByVal 0&, IDC_HAND)
'
'  GUID(0) = &H7BF80980
'  GUID(1) = &H101ABF32
'  GUID(2) = &HAA00BB8B
'  GUID(3) = &HAB0C3000
'
'  With Pic
'    .Size = Len(Pic)
'    .type = vbPicTypeIcon
'    .hBmp = hCur
'    .hpal = 0
'  End With
'
'  Call OleCreatePictureIndirect(Pic, GUID(0), 1, IPic)
'
'  Set GetSystemHandCursor = IPic
'End Function

Private Function GetWindowsDPI() As Double
    Dim hdc As Long, LPX  As Double, LPY As Double
    hdc = GetDC(0)
    LPX = CDbl(GetDeviceCaps(hdc, LOGPIXELSX))
    LPY = CDbl(GetDeviceCaps(hdc, LOGPIXELSY))
    ReleaseDC 0, hdc

    If (LPX = 0) Then
        GetWindowsDPI = 1#
    Else
        GetWindowsDPI = LPX / 96#
    End If
End Function

Private Function gRoundRect(ByVal hGraphics As Long, RECT As RECTL, ByVal Color1 As Long, ByVal Color2 As Long, ByVal Angulo As Single, ByVal BorderColor As Long, ByVal Round As Long, Filled As Boolean) As Long
    Dim hPen As Long
    Dim hBrush As Long
    Dim mPath As Long
    Dim mRound As Long
    
    If m_BorderWidth <> 0 Then GdipCreatePen1 BorderColor, m_BorderWidth * nScale, &H2, hPen   '&H1 * nScale, &H2, hPen
    If Filled Then GdipCreateLineBrushFromRectWithAngleI RECT, Color1, Color2, Angulo + 90, 0, WrapModeTileFlipXY, hBrush
    GdipCreatePath &H0, mPath   '&H0
    
    With RECT
        mRound = GetSafeRound((Round * nScale), .Width * 2, .Height * 2)
        If mRound = 0 Then mRound = 1
            GdipAddPathArcI mPath, .Left, .Top, mRound, mRound, 180, 90
            GdipAddPathArcI mPath, (.Left + .Width) - mRound, .Top, mRound, mRound, 270, 90
            GdipAddPathArcI mPath, (.Left + .Width) - mRound, (.Top + .Height) - mRound, mRound, mRound, 0, 90
            GdipAddPathArcI mPath, .Left, (.Top + .Height) - mRound, mRound, mRound, 90, 90
    End With
    
    GdipClosePathFigures mPath
    GdipFillPath hGraphics, hBrush, mPath
    GdipDrawPath hGraphics, hPen, mPath
    
    Call GdipDeletePath(mPath)
    Call GdipDeleteBrush(hBrush)
    Call GdipDeletePen(hPen)

    gRoundRect = mPath
End Function

'Inicia GDI+
Private Sub InitGDI()
    Dim GdipStartupInput As GDIPlusStartupInput
    GdipStartupInput.GdiPlusVersion = 1&
    Call GdiplusStartup(GdipToken, GdipStartupInput, ByVal 0)
End Sub

Private Function IsMouseOver(hWnd As Long) As Boolean
    Dim Pt As POINTL
    GetCursorPos Pt
    IsMouseOver = (WindowFromPoint(Pt.x, Pt.y) = hWnd)
End Function

'Private Function MousePointerHands(ByVal NewValue As Boolean)
'  If NewValue Then
'    If Ambient.UserMode Then
'      UserControl.MousePointer = vbCustom
'      UserControl.MouseIcon = GetSystemHandCursor
'    End If
'  Else
'    If hCur Then DestroyCursor hCur: hCur = 0
'    UserControl.MousePointer = vbDefault
'    UserControl.MouseIcon = Nothing
'  End If
'
'End Function

Private Function ReadValue(ByVal lProp As Long, Optional Default As Long) As Long
    Dim I       As Long
    For I = 0 To TLS_MINIMUM_AVAILABLE - 1
        If TlsGetValue(I) = lProp Then
            ReadValue = TlsGetValue(I + 1)
            Exit Function
        End If
    Next
    ReadValue = Default
End Function

Private Sub SafeRange(Value, Min, Max)
    If Value < Min Then Value = Min
    If Value > Max Then Value = Max
End Sub

'Termina GDI+
Private Sub TerminateGDI()
    Call GdiplusShutdown(GdipToken)
End Sub

Private Sub tmrEffect_Timer()
If IsMouseOver(UserControl.hWnd) Then
  If m_Opacity < 100 Then
    m_Opacity = m_Opacity + 2
    Refresh
  Else
    Exit Sub
  End If
Else
  m_Opacity = m_InitialOpacity
  Refresh
  tmrEffect.Enabled = False
End If
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
  CopyAmbient
End Sub

Private Sub UserControl_Click()
  RaiseEvent Click
    
End Sub

Private Sub UserControl_Initialize()
    InitGDI
    nScale = GetWindowsDPI
End Sub

'Inicializar propiedades para control de usuario
Private Sub UserControl_InitProperties()
hFontCollection = ReadValue(&HFC)
cl_hWnd = UserControl.ContainerHwnd

  m_Clickable = False
  m_Clicked = False
  Set m_Font = UserControl.Ambient.Font
  Set m_Font2 = UserControl.Ambient.Font
  Set m_Font3 = UserControl.Ambient.Font
  Set m_IconFont = UserControl.Font
  m_CaptionAlignV = eMiddle
  m_CaptionAlignH = eCenter
  m_Caption2AlignV = eMiddle
  m_Caption2AlignH = eCenter
  m_Caption3AlignV = eMiddle
  m_Caption3AlignH = eCenter
  m_Caption = Ambient.DisplayName
  m_Caption2 = Ambient.DisplayName
  m_Caption3 = Ambient.DisplayName
  m_Caption3Enable = False
  
  m_isBoxed = False
  m_IconBox = ebLeft
  m_Filled = True

  m_BorderColor = &HC0&
  m_OldBorderColor = &HC0&
  m_ForeColor = m_def_ForeColor
  m_ForeColor2 = &HFFFFFF
  m_ForeColor3 = &HFFFFFF
  m_Enabled = True
  m_Color1 = m_def_Color1
  m_Color2 = m_def_Color2
  m_Angulo = m_def_Angulo
  m_BorderWidth = 2
  m_CornerCurve = 10
  m_Transparent = True
  m_Opacity = 50
  m_InitialOpacity = m_Opacity
  m_IconCharCode = "&H0"
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
m_Clicked = True
Refresh
RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
tmrEffect.Enabled = True
RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
m_Clicked = False
Refresh
RaiseEvent MouseUp(Button, Shift, x, y)
End Sub

'Cargar valores de propiedad desde el almacén
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

With PropBag
  m_Enabled = .ReadProperty("Enabled", True)
  
  m_Color1 = .ReadProperty("BackColor1", m_def_Color1)
  m_Color2 = .ReadProperty("BackColor2", m_def_Color2)
  m_Angulo = .ReadProperty("BackAngle", m_def_Angulo)
  m_BorderColor = .ReadProperty("BorderColor", &HC0&)
  m_BorderWidth = .ReadProperty("BorderWidth", 1)
  m_CornerCurve = .ReadProperty("CornerCurve", 0)
  m_Filled = .ReadProperty("Filled", True)
  
  m_isBoxed = .ReadProperty("Boxed", False)
  m_IconBox = .ReadProperty("IconBoxSide", ebLeft)
  
  m_ForeColor = .ReadProperty("Caption1Color", m_def_ForeColor)
  m_ForeColor2 = .ReadProperty("Caption2Color", m_def_ForeColor)
  m_ForeColor3 = .ReadProperty("Caption3Color", m_def_ForeColor)
  
  Set m_Font = .ReadProperty("Caption1Font", UserControl.Font)
  m_Caption = .ReadProperty("Caption1", Ambient.DisplayName)
  m_CaptionAlignV = .ReadProperty("Caption1AlignV", 1)
  m_CaptionAlignH = .ReadProperty("Caption1AlignH", 1)
  Set m_Font2 = .ReadProperty("Caption2Font", UserControl.Font)
  m_Caption2 = .ReadProperty("Caption2", Ambient.DisplayName)
  m_Caption2AlignV = .ReadProperty("Caption2AlignV", 1)
  m_Caption2AlignH = .ReadProperty("Caption2AlignH", 1)
  Set m_Font3 = .ReadProperty("Caption3Font", UserControl.Font)
  m_Caption3 = .ReadProperty("Caption3", Ambient.DisplayName)
  m_Caption3AlignV = .ReadProperty("Caption3AlignV", 1)
  m_Caption3AlignH = .ReadProperty("Caption3AlignH", 1)
  m_Caption3Enable = .ReadProperty("Caption3Enable", False)
  
  m_Transparent = .ReadProperty("Transparent", True)
    
  m_BoxedColor = .ReadProperty("BoxedColor", vbWhite)
  m_ChangeBorderOnFocus = .ReadProperty("ChangeColorOnFocus", False)
  m_EffectFade = .ReadProperty("EffectFading", False)
  m_InitialOpacity = .ReadProperty("InitialOpacity", 50)
  
  Set m_IconFont = .ReadProperty("IconFont", UserControl.Font)
  m_IconCharCode = .ReadProperty("IconCharCode", "&H0")
  m_IconForeColor = .ReadProperty("IconForeColor", &H404040)
  m_IconAlignV = .ReadProperty("IconAlignV", 1)
  m_IconAlignH = .ReadProperty("IconAlignH", 1)
  
  m_Clickable = .ReadProperty("Clickable", False)
  
End With
  
  m_Opacity = m_InitialOpacity
  UserControl.Enabled = m_Enabled
  
End Sub

Private Sub UserControl_Resize()
Refresh
End Sub

Private Sub UserControl_Terminate()
TerminateGDI
End Sub

'Escribir valores de propiedad en el almacén
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
  Call .WriteProperty("Enabled", m_Enabled)
  Call .WriteProperty("BackColor1", m_Color1, m_def_Color1)
  Call .WriteProperty("BackColor2", m_Color2, m_def_Color2)
  Call .WriteProperty("BackAngle", m_Angulo, m_def_Angulo)
  Call .WriteProperty("BorderColor", m_BorderColor, &HC0&)
  Call .WriteProperty("BorderWidth", m_BorderWidth, 1)
  Call .WriteProperty("CornerCurve", m_CornerCurve, 0)
  Call .WriteProperty("Filled", m_Filled)
  Call .WriteProperty("Boxed", m_isBoxed)
  Call .WriteProperty("IconBoxSide", m_IconBox)
  
  Call .WriteProperty("Caption1Color", m_ForeColor, m_def_ForeColor)
  Call .WriteProperty("Caption2Color", m_ForeColor2, m_def_ForeColor)
  Call .WriteProperty("Caption3Color", m_ForeColor3, m_def_ForeColor)
  
  Call .WriteProperty("Caption1Font", m_Font, UserControl.Ambient.Font)
  Call .WriteProperty("Caption1", m_Caption, Ambient.DisplayName)
  Call .WriteProperty("Caption1AlignV", m_CaptionAlignV, 1)
  Call .WriteProperty("Caption1AlignH", m_CaptionAlignH, 1)
  Call .WriteProperty("Caption2Font", m_Font2, UserControl.Ambient.Font)
  Call .WriteProperty("Caption2", m_Caption2, Ambient.DisplayName)
  Call .WriteProperty("Caption2AlignV", m_Caption2AlignV, 1)
  Call .WriteProperty("Caption2AlignH", m_Caption2AlignH, 1)
  Call .WriteProperty("Caption3Font", m_Font3, UserControl.Ambient.Font)
  Call .WriteProperty("Caption3", m_Caption3, Ambient.DisplayName)
  Call .WriteProperty("Caption3AlignV", m_Caption3AlignV, 1)
  Call .WriteProperty("Caption3AlignH", m_Caption3AlignH, 1)
  Call .WriteProperty("Caption3Enable", m_Caption3Enable)
  
  Call .WriteProperty("Transparent", m_Transparent, True)
    
  Call .WriteProperty("BoxedColor", m_BoxedColor, vbWhite)
  Call .WriteProperty("ChangeColorOnFocus", m_ChangeBorderOnFocus, False)
  Call .WriteProperty("EffectFading", m_EffectFade, False)
  Call .WriteProperty("InitialOpacity", m_InitialOpacity, 50)
  
  Call .WriteProperty("IconFont", m_IconFont)
  Call .WriteProperty("IconCharCode", m_IconCharCode, 0)
  Call .WriteProperty("IconForeColor", m_IconForeColor, vbButtonText)
  Call .WriteProperty("IconAlignV", m_IconAlignV, 1)
  Call .WriteProperty("IconAlignH", m_IconAlignH, 1)
  
  Call .WriteProperty("Clickable", m_Clickable)
  
End With
  
End Sub

Public Property Get BackAngle() As Single
  BackAngle = m_Angulo
End Property

Public Property Let BackAngle(ByVal New_Angulo As Single)
  m_Angulo = New_Angulo
  PropertyChanged "BackAngle"
  Refresh
End Property

Public Property Get BackColor1() As OLE_COLOR
  BackColor1 = m_Color1
End Property

Public Property Let BackColor1(ByVal New_Color1 As OLE_COLOR)
  m_Color1 = New_Color1
  PropertyChanged "BackColor1"
  Refresh
End Property

Public Property Get BackColor2() As OLE_COLOR
  BackColor2 = m_Color2
End Property

Public Property Let BackColor2(ByVal New_Color2 As OLE_COLOR)
  m_Color2 = New_Color2
  PropertyChanged "BackColor2"
  Refresh
End Property

'Properties-------------------
Public Property Get BorderColor() As OLE_COLOR
  BorderColor = m_BorderColor
End Property

Public Property Let BorderColor(ByVal NewBorderColor As OLE_COLOR)
  m_BorderColor = NewBorderColor
  'm_OldBorderColor = m_BorderColor
  PropertyChanged "BorderColor"
  Refresh
End Property

Public Property Get BoxedColor() As OLE_COLOR
  BoxedColor = m_BoxedColor
End Property

Public Property Let BoxedColor(ByVal NewBoxedColor As OLE_COLOR)
  m_BoxedColor = NewBoxedColor
  PropertyChanged "BoxedColor"
  Refresh
End Property

Public Property Get BorderWidth() As Long
  BorderWidth = m_BorderWidth
End Property

Public Property Let BorderWidth(ByVal NewBorderWidth As Long)
  m_BorderWidth = NewBorderWidth
  PropertyChanged "BorderWidth"
  Refresh
End Property

Public Property Get Caption1Font() As StdFont
  Set Caption1Font = m_Font
End Property

Public Property Set Caption1Font(ByVal New_Font As StdFont)
  Set m_Font = New_Font
  PropertyChanged "Caption1Font"
  Refresh
End Property

Public Property Get Caption1Color() As OLE_COLOR
  Caption1Color = m_ForeColor
End Property

Public Property Let Caption1Color(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor = NewForeColor
  PropertyChanged "Caption1Color"
  Refresh
End Property

Public Property Get Caption1() As String
  Caption1 = m_Caption
End Property

Public Property Let Caption1(ByVal NewCaption As String)
  m_Caption = NewCaption
  PropertyChanged "Caption1"
  Refresh
End Property

Public Property Get Caption1AlignH() As eTextAlignH
  Caption1AlignH = m_CaptionAlignH
End Property

Public Property Let Caption1AlignH(ByVal NewCaptionAlignH As eTextAlignH)
  m_CaptionAlignH = NewCaptionAlignH
  PropertyChanged "Caption1AlignH"
  Refresh
End Property

Public Property Get Caption1AlignV() As eTextAlignV
  Caption1AlignV = m_CaptionAlignV
End Property

Public Property Let Caption1AlignV(ByVal NewCaptionAlignV As eTextAlignV)
  m_CaptionAlignV = NewCaptionAlignV
  PropertyChanged "Caption1AlignV"
  Refresh
End Property

Public Property Get Caption2Font() As StdFont
  Set Caption2Font = m_Font2
End Property

Public Property Set Caption2Font(ByVal New_Font As StdFont)
  Set m_Font2 = New_Font
  PropertyChanged "Caption2Font"
  Refresh
End Property

Public Property Get Caption2Color() As OLE_COLOR
  Caption2Color = m_ForeColor2
End Property

Public Property Let Caption2Color(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor2 = NewForeColor
  PropertyChanged "Caption2Color"
  Refresh
End Property
Public Property Get Caption2() As String
  Caption2 = m_Caption2
End Property

Public Property Let Caption2(ByVal NewCaption As String)
  m_Caption2 = NewCaption
  PropertyChanged "Caption2"
  Refresh
End Property

Public Property Get Caption2AlignH() As eTextAlignH
  Caption2AlignH = m_Caption2AlignH
End Property

Public Property Let Caption2AlignH(ByVal NewCaptionAlignH As eTextAlignH)
  m_Caption2AlignH = NewCaptionAlignH
  PropertyChanged "Caption2AlignH"
  Refresh
End Property

Public Property Get Caption2AlignV() As eTextAlignV
  Caption2AlignV = m_Caption2AlignV
End Property

Public Property Let Caption2AlignV(ByVal NewCaptionAlignV As eTextAlignV)
  m_Caption2AlignV = NewCaptionAlignV
  PropertyChanged "Caption2AlignV"
  Refresh
End Property
'----
Public Property Get Caption3Font() As StdFont
  Set Caption3Font = m_Font3
End Property

Public Property Set Caption3Font(ByVal New_Font As StdFont)
  Set m_Font3 = New_Font
  PropertyChanged "Caption3Font"
  Refresh
End Property

Public Property Get Caption3Color() As OLE_COLOR
  Caption3Color = m_ForeColor3
End Property

Public Property Let Caption3Color(ByVal NewForeColor As OLE_COLOR)
  m_ForeColor3 = NewForeColor
  PropertyChanged "Caption3Color"
  Refresh
End Property

Public Property Get Caption3() As String
  Caption3 = m_Caption3
End Property

Public Property Let Caption3(ByVal NewCaption As String)
  m_Caption3 = NewCaption
  PropertyChanged "Caption3"
  Refresh
End Property

Public Property Get Caption3AlignH() As eTextAlignH
  Caption3AlignH = m_Caption3AlignH
End Property

Public Property Let Caption3AlignH(ByVal NewCaptionAlignH As eTextAlignH)
  m_Caption3AlignH = NewCaptionAlignH
  PropertyChanged "Caption3AlignH"
  Refresh
End Property

Public Property Get Caption3AlignV() As eTextAlignV
  Caption3AlignV = m_Caption3AlignV
End Property

Public Property Let Caption3AlignV(ByVal NewCaptionAlignV As eTextAlignV)
  m_Caption3AlignV = NewCaptionAlignV
  PropertyChanged "Caption3AlignV"
  Refresh
End Property
'----
Public Property Get ChangeBorderOnFocus() As Boolean
  ChangeBorderOnFocus = m_ChangeBorderOnFocus
End Property

Public Property Let ChangeBorderOnFocus(ByVal NewChangeBorderOnFocus As Boolean)
  m_ChangeBorderOnFocus = NewChangeBorderOnFocus
  m_OldBorderColor = m_BorderColor
  PropertyChanged "ChangeBorderOnFocus"
End Property

Public Property Get CornerCurve() As Long
  CornerCurve = m_CornerCurve
End Property

Public Property Let CornerCurve(ByVal NewCornerCurve As Long)
  m_CornerCurve = NewCornerCurve
  PropertyChanged "CornerCurve"
  Refresh
End Property

Public Property Get EffectFading() As Boolean
EffectFading = m_EffectFade
End Property

Public Property Let EffectFading(ByVal vNewValue As Boolean)
m_EffectFade = vNewValue
PropertyChanged "EffectFading"
Refresh
End Property

Public Property Get Enabled() As Boolean
  Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
  m_Enabled = New_Enabled
  PropertyChanged "Enabled"
End Property

Public Property Get Filled() As Boolean
  Filled = m_Filled
End Property

Public Property Let Filled(ByVal New_Filled As Boolean)
  m_Filled = New_Filled
  PropertyChanged "Filled"
  Refresh
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

Public Property Get IconCharCode() As String
    IconCharCode = "&H" & Hex(m_IconCharCode)
End Property

Public Property Let IconCharCode(ByVal New_IconCharCode As String)
    New_IconCharCode = UCase(Replace(New_IconCharCode, Space(1), vbNullString))
    New_IconCharCode = UCase(Replace(New_IconCharCode, "U+", "&H"))
    If Not VBA.Left$(New_IconCharCode, 2) = "&H" And Not IsNumeric(New_IconCharCode) Then
        m_IconCharCode = "&H" & New_IconCharCode
    Else
        m_IconCharCode = New_IconCharCode
    End If
    PropertyChanged "IconCharCode"
    Refresh
End Property

Public Property Get IconFont() As StdFont
    Set IconFont = m_IconFont
End Property

Public Property Set IconFont(New_Font As StdFont)
  Set m_IconFont = New_Font
'    With m_IconFont
'        .Name = New_Font.Name
'        .Size = New_Font.Size
'        .Bold = New_Font.Bold
'        .Italic = New_Font.Italic
'        .Strikethrough = New_Font.Strikethrough
'        .Underline = New_Font.Underline
'        .Weight = New_Font.Weight
'    End With
    PropertyChanged "IconFont"
  Refresh
End Property

Public Property Get IconForeColor() As OLE_COLOR
    IconForeColor = m_IconForeColor
End Property

Public Property Let IconForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_IconForeColor = New_ForeColor
    PropertyChanged "IconForeColor"
    Refresh
End Property

Public Property Get IconAlignH() As eTextAlignH
  IconAlignH = m_IconAlignH
End Property

Public Property Let IconAlignH(ByVal NewIconAlignH As eTextAlignH)
  m_IconAlignH = NewIconAlignH
  PropertyChanged "IconAlignH"
  Refresh
End Property

Public Property Get IconAlignV() As eTextAlignV
  IconAlignV = m_IconAlignV
End Property

Public Property Let IconAlignV(ByVal NewIconAlignV As eTextAlignV)
  m_IconAlignV = NewIconAlignV
  PropertyChanged "IconAlignV"
  Refresh
End Property

Public Property Get InitialOpacity() As Long
  InitialOpacity = m_InitialOpacity
End Property

Public Property Let InitialOpacity(ByVal NewInitialOpacity As Long)
  m_InitialOpacity = NewInitialOpacity
  PropertyChanged "InitialOpacity"
  m_Opacity = m_InitialOpacity
  Refresh
End Property

Public Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property

Public Property Let Transparent(ByVal NewValue As Boolean)
    m_Transparent = NewValue
    PropertyChanged "Transparent"
    Refresh
End Property

Public Property Get Version() As String
Version = App.Major & "." & App.Minor & "." & App.Revision
End Property

Public Property Get Visible() As Boolean
  Visible = Extender.Visible
End Property

Public Property Let Visible(ByVal NewVisible As Boolean)
  Extender.Visible = NewVisible
End Property

Public Property Get Clickable() As Boolean
   Clickable = m_Clickable
End Property

Public Property Let Clickable(ByVal bClickable As Boolean)
   m_Clickable = bClickable
   PropertyChanged "Clickable"
End Property

Public Property Get Boxed() As Boolean
  Boxed = m_isBoxed
End Property

Public Property Let Boxed(ByVal bBoxed As Boolean)
  m_isBoxed = bBoxed
  PropertyChanged "Boxed"
  Refresh
End Property

Public Property Get IconBoxSide() As eBoxed
  IconBoxSide = m_IconBox
End Property

Public Property Let IconBoxSide(ByVal bBoxed As eBoxed)
  m_IconBox = bBoxed
  PropertyChanged "IconBoxSide"
  Refresh
End Property

Public Property Get Caption3Enable() As Boolean
Caption3Enable = m_Caption3Enable
End Property

Public Property Let Caption3Enable(ByVal vNenable As Boolean)
m_Caption3Enable = vNenable
PropertyChanged "Caption3Enable"
Refresh
End Property
