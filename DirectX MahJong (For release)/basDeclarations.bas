Attribute VB_Name = "basDeclarations"
Option Explicit

'////////////////////////////////////////////////////////////////////////////////////
'DirectX
Public dx As New DirectX7
Public dd As DirectDraw7
Public ddsFront As DirectDrawSurface7
Public ddsBack As DirectDrawSurface7
Public ddsBlocks As DirectDrawSurface7
Public ddsBackground As DirectDrawSurface7
Public ddsGameOver As DirectDrawSurface7
Public ddsFadingBlocks As DirectDrawSurface7
'////////////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////////////////////////////
'My Constants
Public Const SCREEN_WIDTH As Integer = 1024
Public Const SCREEN_HEIGHT As Integer = 768
Public Const DISPLAY_DEPTH As Integer = 16

Public Const DEFAULT_PUZZLE_WIDTH As Integer = 20
Public Const DEFAULT_PUZZLE_HEIGHT As Integer = 11

'*** Change NUM_BLOCK_TYPES to a value between 1 and 6 ***
Public Const NUM_BLOCK_TYPES As Integer = 4

Public Const OFFSETX As Integer = 6 'Actually this is the "horizontal" thickness of a block.
Public Const OFFSETY As Integer = 6 'Actually this is the "vertical" thickness of a block

'These two variables determine where the whole MahJongs display start.
Public Const StartX As Integer = 100      '100 pixels to the right from screen edge
Public Const StartY As Integer = 100      '100 pixels down from screen edge

Public Const BLOCK_WIDTH As Integer = 47     'The width of the whole MahJong block
Public Const BLOCK_HEIGHT As Integer = 58    'The height of the whole MahJong block

Public Const FACE_WIDTH As Integer = 41   'The width of the front surface (viewed from above) of MahJong block
Public Const FACE_HEIGHT As Integer = 52  'The height of the front surface (viewed from above) of MahJong block

Public Const COMPLIMENT As String = "Congratulations!! You removed all the blocks!"
'////////////////////////////////////////////////////////////////////////////////////

'////////////////////////////////////////////////////////////////////////////////////
'Types
Public Type POINTAPI
      x As Long
      y As Long
End Type

Public Type BLOCK_INFO
      Index As Integer
      key As String
      BlockType As Integer
      HomeLeft As Integer
      HomeTop As Integer
      XCoord As Integer
      YCoord As Integer
      Left As Long
      Top As Long
      Right As Long
      Bottom As Long
      Exists As Boolean
      HasBeenFound As Boolean
      Fading As Boolean
End Type
'////////////////////////////////////////////////////////////////////////////////////


'////////////////////////////////////////////////////////////////////////////////////
'My variables
Public PuzzleWidth As Integer       'The number of blocks in a row
Public PuzzleHeight As Integer      'The number of blocks in a column

Public TotalBlocks As Integer

Public MovedDistance As Integer
Public FadeStep As Integer

Public boolEndGame As Boolean
Public boolCanRemoveBlocks As Boolean
Public boolGameOver As Boolean
Public boolProcessing As Boolean
Public boolBlocksToDrop As Boolean
Public boolFading As Boolean

'////////////////////////////////////////////////////////////////////////////////////

'Others
Public Board() As Integer

Public Blocks() As BLOCK_INFO

Public BackgroundRect As RECT

Public colSameBlocks As New Collection
Public colExistingBlocks As New Collection

Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
