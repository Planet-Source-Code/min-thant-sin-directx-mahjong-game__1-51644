Attribute VB_Name = "basDroppingAndMovingBlocks"
Option Explicit

Public Sub MoveRightUntilNoMoreBlocks(ByVal Index As Integer, ByVal RightBlock As Integer)
      Dim temp As Integer
      
      Do While Blocks(RightBlock).Exists
            temp = Blocks(RightBlock).Index
            Board(Blocks(RightBlock).XCoord, Blocks(RightBlock).YCoord) = Blocks(Index).Index
            Board(Blocks(RightBlock).XCoord - 1, Blocks(RightBlock).YCoord) = temp
            
            Blocks(RightBlock).XCoord = Blocks(RightBlock).XCoord - 1
            Blocks(Index).XCoord = Blocks(Index).XCoord + 1
            
            Blocks(RightBlock).Left = Blocks(RightBlock).XCoord * FACE_WIDTH + StartX
            Blocks(RightBlock).Right = Blocks(RightBlock).Left + BLOCK_WIDTH
            
            Blocks(Index).Left = Blocks(Index).XCoord * FACE_WIDTH + StartX
            Blocks(Index).Right = Blocks(Index).Left + BLOCK_WIDTH
            
            RightBlock = GetBlockFromCoord(Blocks(Index).XCoord + 1, Blocks(Index).YCoord)
            'Out of range.
            If RightBlock = -1 Then Exit Do
      Loop
End Sub

Public Sub DropBlock(ByVal Index As Integer)
      Dim BottomBlock As Integer
      Dim temp As Integer
      
      BottomBlock = GetBlockFromCoord(Blocks(Index).XCoord, Blocks(Index).YCoord + 1)
      'Out of range. We check this for safety.
      If BottomBlock = -1 Then Exit Sub
      
      temp = Blocks(Index).Index
      Board(Blocks(Index).XCoord, Blocks(Index).YCoord) = Blocks(BottomBlock).Index
      Board(Blocks(Index).XCoord, Blocks(Index).YCoord + 1) = temp
      
      Blocks(Index).YCoord = Blocks(Index).YCoord + 1
      Blocks(BottomBlock).YCoord = Blocks(BottomBlock).YCoord - 1
      
      Blocks(Index).Top = Blocks(Index).YCoord * FACE_HEIGHT + StartY
      Blocks(Index).Bottom = Blocks(Index).Top + BLOCK_HEIGHT
      
      Blocks(BottomBlock).Top = Blocks(BottomBlock).YCoord * FACE_HEIGHT + StartY
      Blocks(BottomBlock).Bottom = Blocks(BottomBlock).Top + BLOCK_HEIGHT
End Sub
