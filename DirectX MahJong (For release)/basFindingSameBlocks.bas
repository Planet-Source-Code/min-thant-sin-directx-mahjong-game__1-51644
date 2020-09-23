Attribute VB_Name = "basFindingSameBlocks"
Option Explicit

 Public Sub FindSameBlocks(ByVal Index As Integer)
      'Left, right, top, and bottom blocks of the block being clicked
      Dim LeftBlock As Integer
      Dim RightBlock As Integer
      Dim TopBlock As Integer
      Dim BottomBlock As Integer
      
      Dim XPos As Integer, YPos As Integer
      
      XPos = Blocks(Index).XCoord
      YPos = Blocks(Index).YCoord
      
      LeftBlock = GetBlockFromCoord(XPos - 1, YPos)
      RightBlock = GetBlockFromCoord(XPos + 1, YPos)
      TopBlock = GetBlockFromCoord(XPos, YPos - 1)
      BottomBlock = GetBlockFromCoord(XPos, YPos + 1)
      
      'Check left block
      If LeftBlock <> -1 Then
            'Don't count it if it has been found.
            If Blocks(LeftBlock).HasBeenFound = False Then
                  'Same letter? (A, B, C, ...)
                  If Blocks(LeftBlock).BlockType = Blocks(Index).BlockType Then
                        'Indicate this block has been found because we don't
                        'want to add it to the collection twice.
                        Blocks(LeftBlock).HasBeenFound = True
                        'Add it to the collection
                        Call colSameBlocks.Add(Blocks(LeftBlock).Index) ', Blocks(LeftBlock).Key)
                        'Find same blocks around this one
                        Call FindSameBlocks(Blocks(LeftBlock).Index)
                  End If
            End If
      End If
      
      'Check right block
      If RightBlock <> -1 Then
            If Blocks(RightBlock).HasBeenFound = False Then
                  If Blocks(RightBlock).BlockType = Blocks(Index).BlockType Then
                        Blocks(RightBlock).HasBeenFound = True
                        Call colSameBlocks.Add(Blocks(RightBlock).Index) ', Blocks(RightBlock).Key)
                        Call FindSameBlocks(Blocks(RightBlock).Index)
                  End If
            End If
      End If
      
      'Check top block (above block)
      If TopBlock <> -1 Then
            If Blocks(TopBlock).HasBeenFound = False Then
                  If Blocks(TopBlock).BlockType = Blocks(Index).BlockType Then
                        Blocks(TopBlock).HasBeenFound = True
                        Call colSameBlocks.Add(Blocks(TopBlock).Index) ', Blocks(TopBlock).Key)
                        Call FindSameBlocks(Blocks(TopBlock).Index)
                  End If
            End If
      End If
      
      'Check bottom block
      If BottomBlock <> -1 Then
            If Blocks(BottomBlock).HasBeenFound = False Then
                  If Blocks(BottomBlock).BlockType = Blocks(Index).BlockType Then
                        Blocks(BottomBlock).HasBeenFound = True
                        Call colSameBlocks.Add(Blocks(BottomBlock).Index) ', Blocks(BottomBlock).Key)
                        Call FindSameBlocks(Blocks(BottomBlock).Index)
                  End If
            End If
      End If
End Sub
