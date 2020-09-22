VERSION 5.00
Begin VB.UserControl ucCanvas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3060
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
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
   LockControls    =   -1  'True
   MousePointer    =   99  'Custom
   ScaleHeight     =   148
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   204
End
Attribute VB_Name = "ucCanvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'================================================
' User control:  ucCanvas.ctl (simplified)
' Author:        Carles P.V.
' Dependencies:  cDIB.cls
' Last revision: 2004.09.15
'================================================

Option Explicit

'-- API:

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const RGN_DIFF As Long = 4

Private Declare Function CreateRectRgn Lib "gdi32" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function FillRgn Lib "gdi32" (ByVal hDC As Long, ByVal hRgn As Long, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TranslateColor Lib "olepro32" Alias "OleTranslateColor" (ByVal Clr As OLE_COLOR, ByVal Palette As Long, Col As Long) As Long

'//

'-- Public Enums.:
Public Enum eWorkModeCts
    [cnvScrollMode]
    [cnvUserMode]
End Enum

'-- Property Variables:
Private m_Zoom      As Long
Private m_WorkMode  As eWorkModeCts
Private m_FitMode   As Boolean
Private m_Enabled   As Boolean
Private m_BackColor As OLE_COLOR

'-- Private Variables:
Private m_Width     As Long
Private m_Height    As Long
Private m_Left      As Long
Private m_Top       As Long
Private m_hPos      As Long
Private m_hMax      As Long
Private m_vPos      As Long
Private m_vMax      As Long
Private m_lsthPos   As Single
Private m_lstvPos   As Single
Private m_lsthMax   As Single
Private m_lstvMax   As Single
Private m_Down      As Boolean
Private m_Pt        As POINTAPI

'-- Event Declarations:
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Public Event MouseDown(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event MouseMove(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event MouseUp(Button As Integer, Shift As Integer, x As Long, y As Long)
Public Event Scroll()
Public Event Resize()

'-- Public objects:
Public DIB As cDIB ' DIB section



'========================================================================================
' UserControl
'========================================================================================

Private Sub UserControl_Initialize()

    '-- Initialize DIB
    Set Me.DIB = New cDIB
    
    '-- Default values
    m_Zoom = 1
    m_WorkMode = [cnvScrollMode]
End Sub

Private Sub UserControl_Terminate()

    '-- Destroy DIB
    Set Me.DIB = Nothing
End Sub

'//

Private Sub UserControl_Resize()

    '-- Resize and refresh
    Call pvResizeCanvas
    Call pvRefreshCanvas
    
    RaiseEvent Resize
End Sub

Private Sub UserControl_Paint()

    '-- Refresh Canvas
    Call pvRefreshCanvas
End Sub

'//

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
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
    
    '-- Mouse down flag / Store values
    m_Down = (Button = vbLeftButton)
    m_Pt.x = x
    m_Pt.y = y
    
    RaiseEvent MouseDown(Button, Shift, pvDIBx(x), pvDIBy(y))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    If (m_Down And m_WorkMode = [cnvScrollMode]) Then
    
        '-- Apply offsets
        m_hPos = m_hPos + (m_Pt.x - x)
        m_vPos = m_vPos + (m_Pt.y - y)
        '-- Check margins
        If (m_hPos < 0) Then m_hPos = 0 Else If (m_hPos > m_hMax) Then m_hPos = m_hMax
        If (m_vPos < 0) Then m_vPos = 0 Else If (m_vPos > m_vMax) Then m_vPos = m_vMax
        '-- Save current position
        m_Pt.x = x
        m_Pt.y = y
        
        If (m_lsthPos <> m_hPos Or m_lstvPos <> m_vPos) Then
            Call pvRefreshCanvas
            RaiseEvent Scroll
        End If
        m_lsthPos = m_hPos
        m_lstvPos = m_vPos
    End If
    
    RaiseEvent MouseMove(Button, Shift, pvDIBx(x), pvDIBy(y))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    '-- Mouse down flag
    m_Down = False
    
    RaiseEvent MouseUp(Button, Shift, pvDIBx(x), pvDIBy(y))
End Sub

'========================================================================================
' Methods
'========================================================================================

Public Sub Repaint()
    Call pvRefreshCanvas
End Sub

Public Sub Resize()
    Call pvResizeCanvas
End Sub

Public Function Scroll(ByVal x As Long, ByVal y As Long) As Boolean

    '-- Apply offsets
    m_hPos = m_hPos - x
    m_vPos = m_vPos - y
    '-- Check margins
    If (m_hPos < 0) Then m_hPos = 0 Else If (m_hPos > m_hMax) Then m_hPos = m_hMax
    If (m_vPos < 0) Then m_vPos = 0 Else If (m_vPos > m_vMax) Then m_vPos = m_vMax
    
    If (m_lsthPos <> m_hPos Or m_lstvPos <> m_vPos) Then
        Call pvRefreshCanvas: Scroll = True
        RaiseEvent Scroll
    End If
    m_lsthPos = m_hPos
    m_lstvPos = m_vPos
End Function

'========================================================================================
' Properties
'========================================================================================

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    Call Me.Repaint
End Property
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled = New_Enabled
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_MemberFlags = "400"
    Enabled = UserControl.Enabled
End Property

Public Property Let FitMode(ByVal New_FitMode As Boolean)
    m_FitMode = New_FitMode
End Property
Public Property Get FitMode() As Boolean
Attribute FitMode.VB_MemberFlags = "400"
    FitMode = m_FitMode
End Property

Public Property Get UserIcon() As StdPicture
Attribute UserIcon.VB_MemberFlags = "400"
    Set UserIcon = UserControl.MouseIcon
End Property
Public Property Set UserIcon(ByVal New_MouseIcon As StdPicture)
    Set UserControl.MouseIcon = New_MouseIcon
    Call pvUpdatePointer
End Property

Public Property Let WorkMode(ByVal New_WorkMode As eWorkModeCts)
    m_WorkMode = New_WorkMode
    Call pvUpdatePointer
End Property
Public Property Get WorkMode() As eWorkModeCts
Attribute WorkMode.VB_MemberFlags = "400"
    WorkMode = m_WorkMode
End Property

Public Property Let Zoom(ByVal New_Zoom As Long)
    m_Zoom = IIf(New_Zoom < 1, 1, New_Zoom)
End Property
Public Property Get Zoom() As Long
Attribute Zoom.VB_MemberFlags = "400"
    Zoom = m_Zoom
End Property

'//

Public Property Get ScaleWidth() As Long
    ScaleWidth = UserControl.ScaleWidth
End Property
Public Property Get ScaleHeight() As Long
    ScaleHeight = UserControl.ScaleHeight
End Property

Public Property Get ScrollHMax() As Long
    ScrollHMax = m_hMax
End Property
Public Property Get ScrollVMax() As Long
    ScrollVMax = m_vMax
End Property
Public Property Get ScrollHPos() As Long
    ScrollHPos = m_hPos
End Property
Public Property Get ScrollVPos() As Long
    ScrollVPos = m_vPos
End Property

'//

Private Sub UserControl_InitProperties()
    m_BackColor = vbApplicationWorkspace
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BackColor = PropBag.ReadProperty("BackColor", vbApplicationWorkspace)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BackColor", m_BackColor, vbApplicationWorkspace)
End Sub

'========================================================================================
' Private
'========================================================================================

Private Sub pvEraseBackground()

  Dim hRgn_1 As Long
  Dim hRgn_2 As Long
  Dim lColor As Long
  Dim hBrush As Long
    
    '-- Create brush (background)
    Call TranslateColor(m_BackColor, 0, lColor)
    hBrush = CreateSolidBrush(lColor)

    '-- Create Cls region (Control Rect. - Canvas Rect.)
    hRgn_1 = CreateRectRgn(0, 0, ScaleWidth, ScaleHeight)
    hRgn_2 = CreateRectRgn(m_Left, m_Top, m_Left + m_Width, m_Top + m_Height)
    Call CombineRgn(hRgn_1, hRgn_1, hRgn_2, RGN_DIFF)
    
    '-- Fill it
    Call FillRgn(hDC, hRgn_1, hBrush)
    
    '-- Clear
    Call DeleteObject(hBrush)
    Call DeleteObject(hRgn_1)
    Call DeleteObject(hRgn_2)
End Sub

Private Sub pvRefreshCanvas()
  
  Dim xOff As Long, yOff As Long
  Dim wDst As Long, hDst As Long
  Dim xSrc As Long, ySrc As Long
  Dim wSrc As Long, hSrc As Long
    
    If (Me.DIB.hDIB <> 0) Then
        
        '-- Get Left and Width of source image rectangle:
        If (m_hMax And Not m_FitMode) Then
            xOff = -m_hPos Mod m_Zoom
            wDst = (m_Width \ m_Zoom) * m_Zoom + 2 * m_Zoom
            xSrc = m_hPos \ m_Zoom
            wSrc = m_Width \ m_Zoom + 2
          Else
            xOff = m_Left
            wDst = m_Width
            xSrc = 0
            wSrc = Me.DIB.Width
        End If
        
        '-- Get Top and Height of source image rectangle:
        If (m_vMax And Not m_FitMode) Then
            yOff = -m_vPos Mod m_Zoom
            hDst = (m_Height \ m_Zoom) * m_Zoom + 2 * m_Zoom
            ySrc = m_vPos \ m_Zoom
            hSrc = m_Height \ m_Zoom + 2
          Else
            yOff = m_Top
            hDst = m_Height
            ySrc = 0
            hSrc = Me.DIB.Height
        End If
        
        '-- Erase background
        Call pvEraseBackground
        '-- Paint visible source rectangle:
        Call Me.DIB.Stretch(hDC, xOff, yOff, wDst, hDst, xSrc, ySrc, wSrc, hSrc)
        
      Else
        '-- Erase background
        Call pvEraseBackground
    End If
End Sub

Private Sub pvResizeCanvas()
    
    With Me.DIB
        
        If (.hDIB <> 0) Then
        
            If (m_FitMode = False) Then
            
                '-- Get new Width
                If (.Width * m_Zoom > ScaleWidth) Then
                    m_hMax = .Width * m_Zoom - ScaleWidth
                    m_Width = ScaleWidth
                  Else
                    m_hMax = 0
                    m_Width = .Width * m_Zoom
                End If
                '-- Get new Height
                If (.Height * m_Zoom > ScaleHeight) Then
                    m_vMax = .Height * m_Zoom - ScaleHeight
                    m_Height = ScaleHeight
                  Else
                    m_vMax = 0
                    m_Height = .Height * m_Zoom
                End If
                '-- Offsets
                m_Left = (ScaleWidth - m_Width) \ 2
                m_Top = (ScaleHeight - m_Height) \ 2
              Else
                Call .GetBestFitInfo(.Width, .Height, ScaleWidth, ScaleHeight, m_Left, m_Top, m_Width, m_Height)
            End If
                                
            '-- Memory position:
            If (m_lsthMax) Then
                m_hPos = (m_lsthPos * m_hMax) \ m_lsthMax
              Else
                m_hPos = m_hMax \ 2
            End If
            If (m_lstvMax) Then
                m_vPos = (m_lstvPos * m_vMax) \ m_lstvMax
              Else
                m_vPos = m_vMax \ 2
            End If
            m_lsthPos = m_hPos: m_lstvPos = m_vPos
            m_lsthMax = m_hMax: m_lstvMax = m_vMax
          
          Else
            '-- 'Hide' canvas
            m_Width = 0: m_Height = 0
        End If
    End With
    
    '-- Update mouse pointer
    Call pvUpdatePointer
End Sub

Private Sub pvUpdatePointer()

    If (m_WorkMode = [cnvScrollMode]) Then
        If ((m_hMax Or m_vMax) And Not m_FitMode) Then
            UserControl.MousePointer = vbSizeAll
          Else
            UserControl.MousePointer = vbDefault
        End If
      Else
        If (Not UserControl.MouseIcon Is Nothing) Then
            UserControl.MousePointer = vbCustom
        End If
    End If
End Sub

Private Function pvDIBx(ByVal x As Long) As Long

    If (Me.DIB.hDIB <> 0) Then
        If (m_FitMode) Then
            pvDIBx = Int((x - m_Left) / (m_Width / Me.DIB.Width))
          Else
            pvDIBx = Int((m_hPos + x - m_Left) / m_Zoom)
        End If
    End If
End Function

Private Function pvDIBy(ByVal y As Long) As Long

    If (Me.DIB.hDIB <> 0) Then
        If (m_FitMode) Then
            pvDIBy = Int((y - m_Top) / (m_Height / Me.DIB.Height))
          Else
            pvDIBy = Int((m_vPos + y - m_Top) / m_Zoom)
        End If
    End If
End Function
