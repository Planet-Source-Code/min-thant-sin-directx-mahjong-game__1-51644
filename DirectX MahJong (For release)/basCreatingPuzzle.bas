Attribute VB_Name = "basCreatingPuzzle"
Option Explicit

Public Sub CreatePuzzle(ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal NumBlockTypes As Integer)
      Dim Index As Integer
      Dim Row As Integer, Col As Integer
      
      PuzzleWidth = nWidth
      PuzzleHeight = nHeight
      
      TotalBlocks = (PuzzleWidth * PuzzleHeight)
      
      ReDim Blocks(0 To TotalBlocks - 1)
      ReDim Board(0 To PuzzleWidth - 1, 0 To PuzzleHeight - 1)
      
      Call CleanUpCollections
      
      Set colSameBlocks = New Collection
                  
      Index = 0
      For Row = 0 To PuzzleHeight - 1
            For Col = (PuzzleWidth - 1) To 0 Step -1
                  Board(Col, Row) = Index
                  
                  With Blocks(Index)
                        .Index = Index
                        .Exists = True
                        .HasBeenFound = False
                        .Fading = False
                        .BlockType = Int(Rnd * NumBlockTypes)
                        .key = Chr(.BlockType + 65) & CStr(.Index)
                        .HomeLeft = (.BlockType * BLOCK_WIDTH)
                        .HomeTop = 0
                        .XCoord = Col
                        .YCoord = Row
                        .Left = Col * (BLOCK_WIDTH - OFFSETX) + 100
                        .Top = Row * (BLOCK_HEIGHT - OFFSETY) + 100
                        .Right = .Left + BLOCK_WIDTH
                        .Bottom = .Top + BLOCK_HEIGHT
                  End With
                                    
                  colExistingBlocks.Add Index, Blocks(Index).key
                  
                  Index = Index + 1
            Next Col
      Next Row
      
      boolGameOver = False
End Sub
