VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H80000000&
   BorderStyle     =   0  'None
   ClientHeight    =   7305
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   8625
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   27.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "MovingNonrectangularImage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   487
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTileMask 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   3375
      Picture         =   "MovingNonrectangularImage.frx":24C6
      ScaleHeight     =   58
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   47
      TabIndex        =   0
      Top             =   450
      Visible         =   0   'False
      Width           =   705
   End
   Begin VB.Label Label1 
      Caption         =   "Remember to set the picTileMask.Visible to False, of course, this label's too!"
      ForeColor       =   &H000000FF&
      Height          =   2505
      Left            =   450
      TabIndex        =   1
      Top             =   1800
      Visible         =   0   'False
      Width           =   7755
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// Thanks to *** Mike Gerwitz *** for his Bouncing Ball code.
'/// You can find his code at :
'/// http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=35597&lngWId=1
'///
'/// I studied his code and used some of his ideas.
'/// I also use some source code from DX SDK.
'/// This is my first game in DirectX.
'/// Graphics are from a game I downloaded from Yahoo!.
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// NOTE : I do NOT attempt to handle errors that may arise while playing the game.
'///              And the code is NOT optimized. Sorry for lack of comments.
'///              I've been busy with college homework. I had to squeeze time out of my free time.
'///              This game was tested on my machine and it worked fine.
'///              Operating System : WindowsXP Home Edition
'///              RAM : 256 MB
'///              CPU Clock : 1.9 GHz
'///              If you have any problems, please forgive me.
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////

'//////////////////////////////////////////////////////////////////////////////////////////////////////////////
'/// DirectX MahJong by Min Thant Sin on Tuesday, February 10, 2004
'/// Feedback, comments, and suggestions are welcome.
'/// Any bugs? Feel free to email me at < minsin999@hotmail.com >
'//////////////////////////////////////////////////////////////////////////////////////////////////////////////

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
      On Error GoTo ErrorHandler
      
      If boolGameOver Then Exit Sub
      
      'All the moving and dropping of the blocks being taken place.
      'Doesn't allow any action while it is being taken place.
      If boolProcessing Then Exit Sub
      
      boolProcessing = True
      
      Const MahJongShape = &H0
      Dim I As Integer, J As Single, K As Integer
      Dim MaskX As Integer, MaskY As Integer
      Dim MouseX As Integer, MouseY As Integer
      Dim BlockClicked As Integer, Index As Integer
      
      Dim NumBlocksClicked(0 To 2) As Integer
      Dim pt As POINTAPI
      
      boolCanRemoveBlocks = False
      
      GetCursorPos pt
      MouseX = pt.x
      MouseY = pt.y
      
      For I = 0 To UBound(NumBlocksClicked)
            NumBlocksClicked(I) = -1
      Next I
      
      J = 0
      For I = 0 To TotalBlocks - 1
            If MouseX >= Blocks(I).Left And MouseX <= Blocks(I).Right Then
                  If MouseY >= Blocks(I).Top And MouseY <= Blocks(I).Bottom Then
                        If Blocks(I).Exists Then
                              MaskX = (MouseX - Blocks(I).Left)
                              MaskY = (MouseY - Blocks(I).Top)
                              
                              If GetPixel(picTileMask.hdc, MaskX, MaskY) = MahJongShape Then
                                    If J >= 3 Then Exit For
                                    NumBlocksClicked(J) = Blocks(I).Index
                                    J = J + 1
                              End If
                        End If
                  End If
            End If
      Next I
      
      If J = 0 Then
      '/// A minor bug has been fixed right here (by including boolProcessing = False )
            boolProcessing = False
            Exit Sub
      End If
      
      Select Case J
      Case 1
            BlockClicked = NumBlocksClicked(0)
      Case 2, 3
            BlockClicked = GetPriorityBlock(NumBlocksClicked())
      End Select
                              
      FindSameBlocks BlockClicked
                        
      If colSameBlocks.Count >= 2 Then
            boolCanRemoveBlocks = True
            boolFading = True
            For I = 1 To colSameBlocks.Count
                  Index = colSameBlocks.Item(I)
                  Blocks(Index).Fading = True
                  colExistingBlocks.Remove Blocks(Index).key
            Next I
            
      Else
            'Blocks(colSameBlocks(1)).HasBeenFound = False
            For I = 1 To colSameBlocks.Count
                  Index = colSameBlocks(I)
                  'Reset found flag to False for next time search
                  Blocks(Index).HasBeenFound = False
            Next I
            boolProcessing = False
      End If
      
      Exit Sub
ErrorHandler:
      Debug.Print "Error occurred in Form_MouseDown : " & Err.Description
      EndGame
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
      If KeyCode = vbKeyEscape Then
            boolEndGame = True
      End If
End Sub
