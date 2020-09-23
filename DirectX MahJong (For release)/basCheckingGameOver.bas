Attribute VB_Name = "basCheckingGameOver"
Option Explicit

Public NumCalled  As Integer

Public Function GameOver() As Boolean
      Dim I As Integer, J As Integer
      Dim LeftBlock As Integer
      Dim RightBlock As Integer
      Dim TopBlock As Integer
      Dim BottomBlock As Integer
            
      NumCalled = NumCalled + 1
      
      Debug.Print "NumCalled : " & NumCalled
      
      GameOver = True
      'Check all the blocks left
      For I = 1 To colExistingBlocks.Count
            J = colExistingBlocks.Item(I)
            LeftBlock = GetBlockFromCoord(Blocks(J).XCoord - 1, Blocks(J).YCoord)
            RightBlock = GetBlockFromCoord(Blocks(J).XCoord + 1, Blocks(J).YCoord)
            BottomBlock = GetBlockFromCoord(Blocks(J).XCoord, Blocks(J).YCoord + 1)
            TopBlock = GetBlockFromCoord(Blocks(J).XCoord, Blocks(J).YCoord - 1)
            
            If LeftBlock <> -1 Then
                  If Blocks(LeftBlock).Exists And Blocks(LeftBlock).BlockType = Blocks(J).BlockType Then
                        GameOver = False
                        Exit Function
                  End If
            End If
            
            If RightBlock <> -1 Then
                  If Blocks(RightBlock).Exists And Blocks(RightBlock).BlockType = Blocks(J).BlockType Then
                        GameOver = False
                        Exit Function
                  End If
            End If
            
            If BottomBlock <> -1 Then
                  If Blocks(BottomBlock).Exists And Blocks(BottomBlock).BlockType = Blocks(J).BlockType Then
                        GameOver = False
                        Exit Function
                  End If
            End If
            
            If TopBlock <> -1 Then
                  If Blocks(TopBlock).Exists And Blocks(TopBlock).BlockType = Blocks(J).BlockType Then
                        GameOver = False
                        Exit Function
                  End If
            End If
      Next I
      
End Function
