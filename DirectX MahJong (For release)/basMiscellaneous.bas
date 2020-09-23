Attribute VB_Name = "basMiscellaneous"
Option Explicit

Public Sub DisplayBackground()
      ddsBack.BltFast 0, 0, ddsBackground, BackgroundRect, DDBLTFAST_WAIT Or DDBLTFAST_NOCOLORKEY
End Sub

Public Sub CleanUpCollections()
      Set colSameBlocks = Nothing
      Set colExistingBlocks = Nothing
End Sub

Public Sub EndGame()
      Set ddsFront = Nothing
      Set ddsBack = Nothing
      Set ddsBlocks = Nothing
      Set ddsFadingBlocks = Nothing
      Set ddsBackground = Nothing
      Set ddsGameOver = Nothing
      
      dd.RestoreDisplayMode
      dd.SetCooperativeLevel 0, DDSCL_NORMAL
    
      Set dd = Nothing
      End
End Sub
