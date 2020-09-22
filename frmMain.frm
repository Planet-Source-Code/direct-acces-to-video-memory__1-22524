VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   6855
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6690
   LinkTopic       =   "Form1"
   ScaleHeight     =   6855
   ScaleWidth      =   6690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DX As DirectX7
Dim DD As DirectDraw7

Dim sPrimary As DirectDrawSurface7
Dim dPrimary As DDSURFACEDESC2

Dim sBackbuffer As DirectDrawSurface7
Dim dBackbuffer As DDSURFACEDESC2

Dim rScreen As RECT
Dim dScreen As DDSURFACEDESC2

Dim Pal(255) As PALETTEENTRY
Dim DDPalette As DirectDrawPalette

Dim bRunning As Boolean

Dim PArray() As Byte
Dim Average As Long

Dim a As Long, b As Long, c As Long, d As Long
Dim x As Long, y As Long, g As Double
Dim cUp As Boolean, dUp As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then bRunning = False
End Sub

Private Sub Form_Load()
    Set DX = New DirectX7
    Set DD = DX.DirectDrawCreate("")
    
    DD.SetCooperativeLevel Me.hWnd, DDSCL_ALLOWMODEX Or DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    DD.SetDisplayMode 640, 480, 8, 0, DDSDM_DEFAULT
    
    dPrimary.lFlags = DDSD_CAPS
    dPrimary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE
    Set sPrimary = DD.CreateSurface(dPrimary)
    
    dBackbuffer.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
    dBackbuffer.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY
    dBackbuffer.lWidth = 640
    dBackbuffer.lHeight = 480
    Set sBackbuffer = DD.CreateSurface(dBackbuffer)
    
    sBackbuffer.SetFontTransparency True
    sBackbuffer.SetForeColor 0
    sPrimary.GetSurfaceDesc dScreen
    
    With rScreen
        .Top = 0
        .Bottom = 0
        .Right = dScreen.lWidth
        .Bottom = dScreen.lHeight
    End With
    
    Randomize Timer
    
    For a = 0 To 255
        Pal(a).red = 0
        Pal(a).green = a
        Pal(a).blue = 0
    Next
        
    Set DDPalette = DD.CreatePalette(DDPCAPS_8BIT Or DDPCAPS_ALLOW256, Pal())
    sPrimary.SetPalette DDPalette
    
    Main
End Sub

Private Sub Main()
    bRunning = True
    
    c = 1
    d = 1
    
    Do Until bRunning = False
        
        sBackbuffer.SetFillColor 0
        For a = 1 To 25
            sBackbuffer.DrawCircle Rnd * 639 + 1, Rnd * 479 + 1, 4
        Next
        sBackbuffer.SetFillColor vbGreen
        For a = 1 To 50
            sBackbuffer.DrawCircle Rnd * 639 + 1, Rnd * 479 + 1, 4
        Next
        sBackbuffer.SetForeColor vbGreen
        sBackbuffer.DrawLine 1, c, 640, c
        sBackbuffer.DrawLine d, 1, d, 480
        If cUp = True Then
            c = c + 2
            If c >= 480 Then cUp = False
        Else
            c = c - 2
            If c <= 0 Then cUp = True
        End If
        If dUp = True Then
            d = d + 2
            If d >= 640 Then dUp = False
        Else
            d = d - 2
            If d <= 0 Then dUp = True
        End If
        
        sBackbuffer.Lock rScreen, dBackbuffer, DDLOCK_WAIT, 0
        sBackbuffer.GetLockedArray PArray()

        For a = 2 To 637
            For b = 2 To 477
                Average = 0
                Average = (Average _
                + PArray(a, b - 1) _
                + PArray(a, b + 1) _
                + PArray(a - 1, b) _
                + PArray(a + 1, b) _
                + PArray(a - 1, b - 1) _
                + PArray(a - 1, b + 1) _
                + PArray(a + 1, b - 1) _
                + PArray(a + 1, b + 1)) _
                / 8 - 1
                If Average < 0 Then Average = 0
                PArray(a, b) = Average
            Next
        Next
        
        g = g + 0.2
        If g > 25.12 Then g = 0
        For a = 1 To 2000
            x = Cos(g + a / (25.12)) * (a / 10) + 320
            y = Sin(g + a / (25.12)) * (a / 10) + 240
            PArray(x, y) = 255
        Next
        For a = 1 To 638
            PArray(a, 0) = 0
            PArray(a, 1) = 0
            PArray(a, 478) = 0
            PArray(a, 479) = 0
        Next
        For b = 0 To 479
            PArray(0, b) = 0
            PArray(1, b) = 0
            PArray(639, b) = 0
            PArray(638, b) = 0
        Next
        
        sBackbuffer.Unlock rScreen
        
        sPrimary.Blt rScreen, sBackbuffer, rScreen, DDBLT_WAIT
        
        DoEvents
    Loop
    
    DD.RestoreAllSurfaces
    DD.RestoreDisplayMode
    Set DD = Nothing
    Set DX = Nothing
    End
End Sub
