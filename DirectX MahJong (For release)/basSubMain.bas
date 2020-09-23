Attribute VB_Name = "basSubMain"
Option Explicit

Public Sub Main()
      'For blitting background.
      With BackgroundRect
            .Left = 0
            .Top = 0
            .Right = .Left + SCREEN_WIDTH
            .Bottom = .Top + SCREEN_HEIGHT
      End With
      
      Load frmMain
            
      InitializeDirectInput
      InitializeDirectDraw
      LaunchGame
End Sub

Sub InitializeDirectDraw()
      On Error GoTo ErrorHandler
    
      Dim ddsdFront As DDSURFACEDESC2
      Dim ddsCaps As DDSCAPS2
      Dim ddsdBlocks As DDSURFACEDESC2
      Dim ddsdFadingBlocks As DDSURFACEDESC2
      Dim ddsdBackground As DDSURFACEDESC2
      Dim ddsdGameOver As DDSURFACEDESC2
      Dim ddKey As DDCOLORKEY
      
      Set dd = dx.DirectDrawCreate("")
      dd.SetCooperativeLevel frmMain.hWnd, DDSCL_EXCLUSIVE Or DDSCL_FULLSCREEN
      dd.SetDisplayMode SCREEN_WIDTH, SCREEN_HEIGHT, DISPLAY_DEPTH, 0, DDSDM_DEFAULT
      
      ddsdFront.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
      ddsdFront.ddsCaps.lCaps = DDSCAPS_COMPLEX Or DDSCAPS_FLIP Or DDSCAPS_PRIMARYSURFACE
      ddsdFront.lBackBufferCount = 1
      Set ddsFront = dd.CreateSurface(ddsdFront)
      
      'Back buffer
      ddsCaps.lCaps = DDSCAPS_BACKBUFFER
      Set ddsBack = ddsFront.GetAttachedSurface(ddsCaps)
      
      'Blocks, background, etc surfaces
      Set ddsBlocks = dd.CreateSurfaceFromFile(App.Path & "\Graphics\MahJongTiles.bmp", ddsdBlocks)
      Set ddsFadingBlocks = dd.CreateSurfaceFromFile(App.Path & "\Graphics\AlphaFadingBlocks.bmp", ddsdFadingBlocks)
      Set ddsBackground = dd.CreateSurfaceFromFile(App.Path & "\Graphics\Background.bmp", ddsdBackground)
      Set ddsGameOver = dd.CreateSurfaceFromFile(App.Path & "\Graphics\gameover.bmp", ddsdGameOver)
      
      'Black
      ddKey.low = 0
      ddKey.high = 0
      
      ddsBlocks.SetColorKey DDCKEY_SRCBLT, ddKey
      ddsFadingBlocks.SetColorKey DDCKEY_SRCBLT, ddKey
      ddsGameOver.SetColorKey DDCKEY_SRCBLT, ddKey
      
      Exit Sub
ErrorHandler:
      Debug.Print Err.Description
      EndGame
End Sub

'//////////////////////////////////////////////////////////////////////////////
'This is the main game loop
'//////////////////////////////////////////////////////////////////////////////
Sub LaunchGame()
      On Error GoTo ErrorHandler
      
      Dim I As Integer, J As Integer
      Dim Index As Integer
      Dim TargetTick As Long
      Dim EmptyRect As RECT
            
      boolEndGame = False
      boolCanRemoveBlocks = False
      boolGameOver = False
      boolProcessing = False
      boolBlocksToDrop = False
      boolFading = False
      
      'Create puzzle with a 17x11 dimension, and 6 block types
      CreatePuzzle DEFAULT_PUZZLE_WIDTH, DEFAULT_PUZZLE_HEIGHT, NUM_BLOCK_TYPES
      
      Do
            TargetTick = dx.TickCount
            
            GetInput
            
            ddsBack.BltColorFill EmptyRect, 0
            DisplayBackground
            
            If boolCanRemoveBlocks Then
                  If boolFading Then
                        Call FadeMahJongs
                  Else
                        Call DropMahJongs
                        DoEvents
                        
                        If boolBlocksToDrop = False Then
                              Call ShiftMahJongs
                              boolCanRemoveBlocks = False
                        End If
                  End If
            End If
            
            DisplayMahJongs
                        
            'In case the animation is too fast, put the following code
            'Do Until dx.TickCount - TargetTick > 2
            'Loop
            
            ddsFront.Flip Nothing, DDFLIP_WAIT
            DoEvents
      Loop Until boolEndGame
      
      EndGame
      Exit Sub
ErrorHandler:
      Debug.Print "Error occurred in LaunchGame sub : " & Err.Description
      EndGame
End Sub
