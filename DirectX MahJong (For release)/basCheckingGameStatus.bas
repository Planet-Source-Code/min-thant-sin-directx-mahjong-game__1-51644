Attribute VB_Name = "basCheckingGameStatus"
Option Explicit

Sub CheckGameStatus()
      'Check if it's game over.
      If Not GameOver Then Exit Sub
      
      boolGameOver = True
      
      Dim KeyboardState(0 To 255) As Byte
      Dim GameOverRect As RECT
      
      DisplayMahJongs
      
      'Game Over graphics.
      With GameOverRect
            .Left = 0
            .Top = 0
            .Right = .Left + 672
            .Bottom = .Top + 263
      End With
      
      'Center the graphic and blit it.
      ddsBack.BltFast (SCREEN_WIDTH - 672) \ 2, (SCREEN_HEIGHT - 263) \ 2, _
      ddsGameOver, GameOverRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                              
      'If no blocks left, that is, the player removed all the MahJongs,
      'display compliment text "Congratulations! You removed all the blocks!"
      If colExistingBlocks.Count = 0 Then
            ddsBack.SetForeColor frmMain.ForeColor
            ddsBack.SetFont frmMain.Font
            ddsBack.DrawText (SCREEN_WIDTH - frmMain.TextWidth(COMPLIMENT)) \ 2, SCREEN_HEIGHT - frmMain.TextHeight(COMPLIMENT), COMPLIMENT, False
      End If
      
      ddsFront.Flip Nothing, 0
            
      'Ask the user to press either Enter or Esc key.
      Do
            diKeyBoard.GetDeviceState 256, KeyboardState(0)
            DoEvents
      Loop Until ((KeyboardState(DIK_RETURN) And &H80) <> 0) Or _
                       ((KeyboardState(DIK_ESCAPE) And &H80) <> 0)
      
      'The player presses Esc key
      If (KeyboardState(DIK_ESCAPE) And &H80) <> 0 Then
            EndGame
      Else  'Enter key.
            CreatePuzzle DEFAULT_PUZZLE_WIDTH, DEFAULT_PUZZLE_HEIGHT, NUM_BLOCK_TYPES
      End If
End Sub
