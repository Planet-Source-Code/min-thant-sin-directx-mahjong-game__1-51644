Attribute VB_Name = "basManipulatingBlocks"
Option Explicit

Sub DisplayMahJongs()
      Dim Row As Integer, Col As Integer
      Dim Index As Integer
      Dim BlockRect As RECT
      
      Index = 0
      
      For Col = (PuzzleWidth - 1) To 0 Step -1
            For Row = 0 To (PuzzleHeight - 1)
                  Index = Board(Col, Row)
                  
                  If Blocks(Index).Exists Then
                        With BlockRect
                              .Left = Blocks(Index).HomeLeft
                              .Top = Blocks(Index).HomeTop
                              .Right = .Left + BLOCK_WIDTH
                              .Bottom = .Top + BLOCK_HEIGHT
                        End With
                        
                        If Blocks(Index).Fading Then
                              ddsBack.BltFast Blocks(Index).Left, Blocks(Index).Top, _
                              ddsFadingBlocks, BlockRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                        Else
                              ddsBack.BltFast Blocks(Index).Left, Blocks(Index).Top, _
                              ddsBlocks, BlockRect, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
                        End If
                  End If
                  
                  DoEvents
            Next Row
      Next Col
End Sub

Sub ShiftMahJongs()
      '///////////////////////////////////////////////////////////
      'Moving the blocks that are behind the gap to the left
      '///////////////////////////////////////////////////////////
      
      Dim RightBlock As Integer     'Blank block's right block
      Dim BaseBlock As Integer      'Bottom-most block (if this block is blank, this indicates a gap)
      Dim BlankBlock As Integer     'Blank block we're checking from left to right direction
      Dim OldXCoord As Integer     'The base block's X Coord
      Dim Row As Integer, Col As Integer  'Row and column
      Dim colBlocksToMove As New Collection     'Blocks to shift to the left
      Dim I As Integer
      
      'We check from PuzzleWidth - 2 because PuzzleWidth - 1 is
      'the last column and there can be no blocks after this.
      For I = (PuzzleWidth - 2) To 0 Step -1
            'Reset Collection
            Set colBlocksToMove = Nothing
            Set colBlocksToMove = New Collection

            'Base (bottom-most) block
            BaseBlock = Board(I, PuzzleHeight - 1)
            'False means it is a gap.
            If Blocks(BaseBlock).Exists = False Then
                  'Mark this blank block's X coordinate
                  OldXCoord = Blocks(BaseBlock).XCoord

                  'Check from bottom to top and...
                  For Row = (PuzzleHeight - 1) To 0 Step -1
                        'Left to right direction
                        For Col = 0 To (PuzzleWidth - 1)
                              BlankBlock = GetBlockFromCoord(Col, Row)
                              'If this block is blank and its XCoord is equal or
                              'greater than that of base block...
                              If Blocks(BlankBlock).Exists = False And Blocks(BlankBlock).XCoord >= OldXCoord Then
                                    RightBlock = GetBlockFromCoord(Blocks(BlankBlock).XCoord + 1, Blocks(BlankBlock).YCoord)
                                    '-1 means this block is at the last column and there is no block to exchange.
                                    If RightBlock <> -1 Then
                                          'If the right block exists, exchange it with the blank block.
                                          If Blocks(RightBlock).Exists Then
                                                Call MoveRightUntilNoMoreBlocks(Blocks(BlankBlock).Index, Blocks(RightBlock).Index)
                                          End If
                                    End If
                              End If
                        Next Col
                  Next Row
            End If
            DoEvents
      Next I
      
      CheckGameStatus
      boolProcessing = False
End Sub

Sub DropMahJongs()
      Dim I As Integer, K As Integer
      Dim Index As Integer
      Dim AboveBlock As Integer
      Dim tmpTop As Single
      Dim colBlocksToDrop As New Collection
      
      boolBlocksToDrop = False
      
      'Check all the blank blocks
      For I = 1 To colSameBlocks.Count
            'Retrieve index from collection
            Index = colSameBlocks(I)
            'Get the above block's index
            AboveBlock = GetBlockFromCoord(Blocks(Index).XCoord, Blocks(Index).YCoord - 1)
            If AboveBlock <> -1 Then
                  If Blocks(AboveBlock).Exists Then
                        boolBlocksToDrop = True
                        colBlocksToDrop.Add AboveBlock
                  End If
            End If
      Next I
      
      Dim StepVal As Integer
      'Divide it by larger value to slow the animation down, or
      'divide it by smaller value to speed up animation.
      StepVal = FACE_HEIGHT \ 10
      
      If boolBlocksToDrop Then
            For I = 1 To colBlocksToDrop.Count
                  K = colBlocksToDrop.Item(I)
                  With Blocks(K)
                        .Top = .Top + StepVal
                  End With
            Next I
                  
            'Dropped distance.
            MovedDistance = MovedDistance + StepVal
                  
            If MovedDistance >= FACE_HEIGHT Then
                  MovedDistance = 0
                  For I = 1 To colBlocksToDrop.Count
                        K = colBlocksToDrop.Item(I)
                        DropBlock K
                  Next I
            End If
                  
            Set colBlocksToDrop = Nothing
            Set colBlocksToDrop = New Collection
      Else
            Set colSameBlocks = Nothing
            Set colSameBlocks = New Collection
      End If
      
End Sub

Sub FadeMahJongs()
      Dim I As Integer
      Dim Index As Integer
      
      For I = 1 To colSameBlocks.Count
            Index = colSameBlocks(I)
            Blocks(Index).HomeTop = (FadeStep * BLOCK_HEIGHT)
      Next I
      
      FadeStep = FadeStep + 1
      
      '19 animation frames
      If FadeStep >= 19 Then
            FadeStep = 0
            boolFading = False
            
            For I = 1 To colSameBlocks.Count
                  Index = colSameBlocks(I)
                  Blocks(Index).Exists = False
                  Blocks(Index).Fading = False
            Next I
      End If
End Sub
