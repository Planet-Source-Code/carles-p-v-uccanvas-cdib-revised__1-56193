VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form fTest 
   Caption         =   "ucCanvas - Test"
   ClientHeight    =   5505
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   6840
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5505
   ScaleWidth      =   6840
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6240
      Top             =   135
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      Filter          =   "Images (*.bmp;*.jpeg;*.gif)|*.bmp;*.jpg;*.jpeg;*.gif"
   End
   Begin Test.ucCanvas ucCanvas1 
      Height          =   5205
      Left            =   75
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   135
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   9181
   End
   Begin VB.Menu mnuFileTop 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Load image..."
         Index           =   0
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Save (test)"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Exit"
         Index           =   3
      End
   End
   Begin VB.Menu mnuOptionsTop 
      Caption         =   "&Options"
      Begin VB.Menu mnuOptions 
         Caption         =   "&Force 32-bit load "
         Index           =   0
      End
   End
   Begin VB.Menu mnuZoomTop 
      Caption         =   "&Zoom"
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom &in"
         Index           =   0
      End
      Begin VB.Menu mnuZoom 
         Caption         =   "Zoom &out"
         Index           =   1
      End
   End
   Begin VB.Menu mnuDitherTop 
      Caption         =   "API &dither"
      Begin VB.Menu mnuDither 
         Caption         =   "&Black and white"
         Index           =   0
      End
      Begin VB.Menu mnuDither 
         Caption         =   "&VGA"
         Index           =   1
      End
      Begin VB.Menu mnuDither 
         Caption         =   "&Halftone-216"
         Index           =   2
      End
   End
   Begin VB.Menu mnuResampleTop 
      Caption         =   "API &resample"
      Begin VB.Menu mnuResample 
         Caption         =   "Thumbnail (24-bit &100x100 max)"
         Index           =   0
      End
      Begin VB.Menu mnuResample 
         Caption         =   "Thumbnail (24-bit &200x200 max)"
         Index           =   1
      End
      Begin VB.Menu mnuResample 
         Caption         =   "Thumbnail (24-bit &300x300 max)"
         Index           =   2
      End
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-- API:

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long

'-- Private constants:

Private Const PALETTE002 As String = _
    "000000FFFFFF"
    
Private Const PALETTE016 As String = _
    "000000000080008000008080800000800080808000C0C0C08080800000FF00FF0000FFFFFF0000FF00FFFFFF00FFFFFF"
    
Private Const PALETTE256 As String = _
    "000000330000660000990000CC0000FF0000003300333300663300993300CC3300FF3300006600336600666600996600" & _
    "CC6600FF6600009900339900669900999900CC9900FF990000CC0033CC0066CC0099CC00CCCC00FFCC0000FF0033FF00" & _
    "66FF0099FF00CCFF00FFFF00000033330033660033990033CC0033FF0033003333333333663333993333CC3333FF3333" & _
    "006633336633666633996633CC6633FF6633009933339933669933999933CC9933FF993300CC3333CC3366CC3399CC33" & _
    "CCCC33FFCC3300FF3333FF3366FF3399FF33CCFF33FFFF33000066330066660066990066CC0066FF0066003366333366" & _
    "663366993366CC3366FF3366006666336666666666996666CC6666FF6666009966339966669966999966CC9966FF9966" & _
    "00CC6633CC6666CC6699CC66CCCC66FFCC6600FF6633FF6666FF6699FF66CCFF66FFFF66000099330099660099990099" & _
    "CC0099FF0099003399333399663399993399CC3399FF3399006699336699666699996699CC6699FF6699009999339999" & _
    "669999999999CC9999FF999900CC9933CC9966CC9999CC99CCCC99FFCC9900FF9933FF9966FF9999FF99CCFF99FFFF99" & _
    "0000CC3300CC6600CC9900CCCC00CCFF00CC0033CC3333CC6633CC9933CCCC33CCFF33CC0066CC3366CC6666CC9966CC" & _
    "CC66CCFF66CC0099CC3399CC6699CC9999CCCC99CCFF99CC00CCCC33CCCC66CCCC99CCCCCCCCCCFFCCCC00FFCC33FFCC" & _
    "66FFCC99FFCCCCFFCCFFFFCC0000FF3300FF6600FF9900FFCC00FFFF00FF0033FF3333FF6633FF9933FFCC33FFFF33FF" & _
    "0066FF3366FF6666FF9966FFCC66FFFF66FF0099FF3399FF6699FF9999FFCC99FFFF99FF00CCFF33CCFF66CCFF99CCFF" & _
    "CCCCFFFFCCFF00FFFF33FFFF66FFFF99FFFFCCFFFFFFFFFF000000000000000000000000000000000000000000000000" & _
    "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000" & _
    "000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000000"

'-- Private variables:

Private m_oDIBBuffer As cDIB





Private Sub Form_Load()
    
    Set m_oDIBBuffer = New cDIB
    
    mnuZoom(0).Caption = mnuZoom(0).Caption & vbTab & "[+]"
    mnuZoom(1).Caption = mnuZoom(1).Caption & vbTab & "[-]"
    
    mnuDitherTop.Enabled = pvIsWindowsNT
    mnuDither(0).Caption = mnuDither(0).Caption & vbTab & "1-bpp"
    mnuDither(1).Caption = mnuDither(1).Caption & vbTab & "4-bpp"
    mnuDither(2).Caption = mnuDither(2).Caption & vbTab & "8-bpp"
    
    mnuResampleTop.Enabled = pvIsWindowsNT
End Sub

Private Sub Form_Resize()

    Call ucCanvas1.Move(0, 0, Me.ScaleWidth, Me.ScaleHeight)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyAdd:      Call mnuZoom_Click(0)
        Case vbKeySubtract: Call mnuZoom_Click(1)
    End Select
End Sub





Private Sub mnuFile_Click(Index As Integer)
    
    On Error GoTo errH
    
    Select Case Index
    
        Case 0 '-- Load image...
                
            Call CommonDialog1.ShowOpen
            DoEvents
            
            With ucCanvas1
            
                Call .DIB.CreateFromStdPicture(VB.LoadPicture(CommonDialog1.Filename), Force32bpp:=mnuOptions(0).Checked)
                Call .Resize
                Call .Repaint
                
                Call .DIB.CloneTo(m_oDIBBuffer)
                Debug.Print "Image: " & .DIB.Width & "x" & .DIB.Height & "x" & .DIB.BPP & "bpp"
            End With
            
        Case 1 '-- Save (test)
        
            Call ucCanvas1.DIB.Save(App.Path & "\Test.bmp")
            
        Case 3 '-- Exit
        
            Call Unload(Me)
    End Select
    Exit Sub
    
errH:
    Debug.Print "Error!: "; Err.Description & "."
End Sub

Private Sub mnuOptions_Click(Index As Integer)
    
    mnuOptions(0).Checked = Not mnuOptions(0).Checked
End Sub

Private Sub mnuZoom_Click(Index As Integer)
    
    With ucCanvas1
        
        If (.DIB.hDIB <> 0) Then
            
            Select Case Index
                Case 0 '-- Zoom in
                    If (.Zoom < 15) Then .Zoom = .Zoom + 1
                Case 1 '-- Zoom out
                    If (.Zoom > 1) Then .Zoom = .Zoom - 1
            End Select
            
            Call .Resize
            Call .Repaint
        End If
    End With
End Sub

Private Sub mnuDither_Click(Index As Integer)
       
    With ucCanvas1
        
        '-- Create DIB and set palette
        With .DIB
        
            Select Case Index
            
                Case 0 '-- 1-bpp
                
                    Call .Create(m_oDIBBuffer.Width, m_oDIBBuffer.Height, [01_bpp])
                    Call .SetPalette(pvExtractPalette(PALETTE002))
                
                Case 1 '-- 4-bpp
                
                    Call .Create(m_oDIBBuffer.Width, m_oDIBBuffer.Height, [04_bpp])
                    Call .SetPalette(pvExtractPalette(PALETTE016))
                
                Case 2 '-- 8-bpp
                
                    Call .Create(m_oDIBBuffer.Width, m_oDIBBuffer.Height, [08_bpp])
                    Call .SetPalette(pvExtractPalette(PALETTE256))
            End Select
        End With
        
        '-- Resize canvas
        Call .Resize
        
        '-- 'API dither' (Ordered dither*)
        '   *Use SetBrushOrgEx() to set 'brush' origin ('dither matrix' in this case)
        With .DIB
            Call m_oDIBBuffer.Paint(.hDC, , , , [sbmHalftone])
        End With
    
        '-- Refresh
        Call .Repaint
    End With
End Sub

Private Sub mnuResample_Click(Index As Integer)

  Dim lBFx    As Long, lBFW As Long
  Dim lBFy    As Long, lBFH As Long
  Dim lFitMax As Long
 
    With ucCanvas1
    
        '-- Create thumbnail DIB
        lFitMax = 100 * (Index + 1)
        With .DIB
            Call .GetBestFitInfo(m_oDIBBuffer.Width, m_oDIBBuffer.Height, lFitMax, lFitMax, lBFx, lBFy, lBFW, lBFH)
            Call .Create(lBFW, lBFH, [24_bpp])
        End With
        
        '-- Resize canvas
        Call .Resize
        
        '-- 'API resample' (Interpolated)
        With .DIB
            Call m_oDIBBuffer.Stretch(.hDC, 0, 0, lBFW, lBFH, , , , , , [sbmHalftone])
        End With
        
        '-- Refresh
        Call .Repaint
    End With
End Sub





' = Private =============================================================================

Private Function pvExtractPalette(sHEXPalette As String) As Byte()
  
  Dim aPal() As Byte
  Dim lEnts  As Long
  Dim lEnt   As Long
  Dim lEntQ  As Long
  Dim lClrQ  As Long
    
    lEnts = Len(sHEXPalette) \ 6: ReDim aPal(0 To 4 * lEnts - 1)
      
    For lEnt = 0 To lEnts - 1
    
        lEntQ = 4 * lEnt
        lClrQ = CLng("&H" & Mid$(sHEXPalette, (lEnt * 6) + 1, 6))
        
        aPal(lEntQ + 0) = (lClrQ And &HFF&)
        aPal(lEntQ + 1) = (lClrQ And &HFF00&) \ &H100&
        aPal(lEntQ + 2) = (lClrQ And &HFF0000) \ &H10000
    Next lEnt
    
    pvExtractPalette = aPal()
End Function

Private Function pvIsWindowsNT() As Boolean

  Dim uOSVI As OSVERSIONINFO
  
    uOSVI.dwOSVersionInfoSize = Len(uOSVI)
   
    If (GetVersionEx(uOSVI)) Then
        pvIsWindowsNT = (uOSVI.dwPlatformId = 2)
    End If
End Function
