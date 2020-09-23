Attribute VB_Name = "basRetrievingBlockInformation"
Option Explicit

Public Function GetPriorityBlock(ByRef BlocksArray() As Integer) As Integer
      Dim MinX As Integer, MaxY As Integer
      Dim I As Integer, J As Integer
      
      'The block which has the largest YCoord & smallest XCoord gets priority
      GetPriorityBlock = -1
      
      J = BlocksArray(0)
      MinX = Blocks(J).XCoord
      MaxY = Blocks(J).YCoord
      
      For I = 1 To UBound(BlocksArray())
            If BlocksArray(I) <> -1 Then
                  If Blocks(BlocksArray(I)).XCoord < MinX Then
                        J = BlocksArray(I)
                        MinX = Blocks(J).XCoord
                  End If
                  
                  If Blocks(BlocksArray(I)).YCoord > MaxY Then
                        J = BlocksArray(I)
                        MaxY = Blocks(J).YCoord
                  End If
            End If
      Next I
      
      GetPriorityBlock = J
End Function

Public Function GetBlockFromCoord(ByVal XCoord As Integer, ByVal YCoord As Integer) As Integer
      GetBlockFromCoord = -1
      
      If XCoord >= 0 And XCoord <= PuzzleWidth - 1 Then
            If YCoord >= 0 And YCoord <= PuzzleHeight - 1 Then
                  GetBlockFromCoord = Board(XCoord, YCoord)
            End If
      End If
End Function
