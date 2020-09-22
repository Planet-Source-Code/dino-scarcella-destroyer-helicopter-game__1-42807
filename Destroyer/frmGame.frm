VERSION 5.00
Begin VB.Form frmGame 
   Caption         =   "Destroyer"
   ClientHeight    =   5460
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   364
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   532
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrHeli 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'The following API calls are for:

'blitting
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

'code timer
Private Declare Function GetTickCount Lib "kernel32" () As Long

'creating buffers / loading sprites
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long

'loading sprites
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long

'cleanup
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

'end of API here...

'our Buffer's DC
Public myBackBuffer As Long
Public myBufferBMP As Long

'The DC (Device Context) of our sprite/graphic
Public myBuildingSprite, myHeliSprite, myMissileSprite, myExplosion1Sprite, myExplosion2Sprite As Long

'player type to write the highest score to the file score.log
Private Type player
 name As String * 15
 score As Integer
End Type
Private BlankPlayer As player

'coordinates of our sprite/graphic on the screen
Public heliX, heliY, missileX, missileY As Long

'height of each structure;speed of helicopter at this point(once one has destroyed
'all structures, the speed increases); current points ;i and j are just counters
'for the for loops
Dim structureheight(), speed, points, i, j As Integer

'one of the rules of this game is that once you have shot a missile,you can not
'shoot another one until it has hit a building or it has dissappeared from the
'screen. This is a boolean determining if you can shoot a missile or not
Public freetoshoot As Boolean

Public Function LoadGraphicDC(sFileName As String) As Long
'cheap error handling
On Error Resume Next

'temp variable to hold our DC address
Dim LoadGraphicDCTEMP As Long

'create the DC address compatible with
'the DC of the screen
LoadGraphicDCTEMP = CreateCompatibleDC(GetDC(0))

'load the graphic file into the DC...
SelectObject LoadGraphicDCTEMP, LoadPicture(sFileName)

'return the address of the file
LoadGraphicDC = LoadGraphicDCTEMP
End Function

Private Sub Form_Load()
'create a compatable DC for the back buffer..
myBackBuffer = CreateCompatibleDC(GetDC(0))

'create a compatible bitmap surface for the DC
'that is the size of our form.. (532 X 364)
'NOTE - the bitmap will act as the actual graphics surface inside the DC
'because without a bitmap in the DC, the DC cannot hold graphical data..
myBufferBMP = CreateCompatibleBitmap(GetDC(0), 532, 364)

'final step of making the back buffer...
'load our created blank bitmap surface into our buffer
'(this will be used as our canvas to draw-on off screen)
SelectObject myBackBuffer, myBufferBMP

'before we can blit to the buffer, we should fill it with black
BitBlt myBackBuffer, 0, 0, 532, 364, 0, 0, 0, vbWhiteness

'load our sprite (using the function we made)
myBuildingSprite = LoadGraphicDC(App.Path & "\building.bmp")
myMissileSprite = LoadGraphicDC(App.Path & "\missile.gif")
myHeliSprite = LoadGraphicDC(App.Path & "\helicopter.gif")
myExplosion1Sprite = LoadGraphicDC(App.Path & "\explosion1.bmp")
myExplosion2Sprite = LoadGraphicDC(App.Path & "\explosion2.bmp")

'CreateRandomBuilding fills our array of structureheight() with random values so
'that our buildings are each with a random height
Call CreateRandomBuilding

'Our building is painted according to the random values we have assigned in
'CreateRandomBuilding
Call paintBuilding

'freetoshot is made true so that we are able to shoot when the game starts
freetoshoot = True

'the X co-ordinate of the helicopter is put off the screen so that we can gradually
'move the helicopter into the screen from left to right. The Y co-ordinate of the
'helicopter is started at the top.
heliX = -48
heliY = 0

'the speed that the helicopter travels is initialized
speed = 3

'our points are reset for the new game and put in the caption of the form as zero
points = 0
frmGame.Caption = "--Destroyer--  Points: 0"

'our helicopter is going to move at regular intervals, so we need to enable our
'timer now
tmrHeli.Enabled = True
End Sub

Private Sub CreateRandomBuilding()
'this function creates the array holding the heights of each structure, we start at
'one because you can see in the game that the first space is left out. This is done
'to give the player a chance to shoot the structure on the left with a horizontal
'rocket when they move down a level.
 For i = 1 To 17
  Randomize
  ReDim Preserve structureheight(i)
  'we create a building with a base of three, added to a random number from 0 to 9.
  'Therefore each building's height ranges from 3 to 12
  structureheight(i) = 3 + Int(10 * Rnd())
 Next i
End Sub

Private Sub paintBuilding()
 'this nested for loop goes through each building (17 buildings), and builds each
 'building from the bottom up according to it's height
 For i = 1 To 17
  'if a building has been bombed entirely,then don't draw it, and go to the next
  'building
  If structureheight(i) = 0 Then GoTo skipbuilding
  
  For j = 1 To structureheight(i)
   'blit sprites to the back-buffer. Paint the buildings in multiples of 28 pixels
   'from left to right (28 * i). And paint the individual structures of the
   'buildings in multiples of 28 pixels from bottom to top (frmGame.ScaleHeight - (28 * j))
   BitBlt myBackBuffer, 28 * i, frmGame.ScaleHeight - (28 * j), 28, 28, myBuildingSprite, 0, 0, vbSrcPaint
  Next j
  
skipbuilding:
 Next i
 
 'now blit the backbuffer to the form...
 BitBlt Me.hdc, 0, 0, 538, 368, myBackBuffer, 0, 0, vbSrcCopy
End Sub

Private Sub tmrHeli_Timer() 'this timer moves our helicopter at regular intervals
  'when our helicopter has destroyed all structures,we want to know this so that
  'we can move to the top again. Therefore we use the boolean structuresleft.
  Dim structuresleft As Boolean
  
  'clear the place where the helicopter used to be (we do this by filling in the
  'old sprites place with black). Speed is the intervals which it moves at on the
  'screen, that's why we place the black at heliX - speed
  BitBlt myBackBuffer, heliX - speed, heliY, 48, 28, 0, 0, 0, vbBlackness

  'blit the helicopter to the back-buffer at co-ordinates heliX,heliY
  BitBlt myBackBuffer, heliX, heliY, 48, 28, myHeliSprite, 0, 0, vbSrcPaint
  
  'now blit the backbuffer to the form...
  BitBlt Me.hdc, 0, 0, 532, 364, myBackBuffer, 0, 0, vbSrcCopy
  
  'increment heliX at the speed that our helicopter is moving
  heliX = heliX + speed
  
  'if the helicopter has gone over the side of the form then start it on the left
  'again. And move the helicopter down by it's height. Note the image size of the
  'helicopter is 48x28.
  If heliX >= 532 Then
   heliX = -48
   heliY = heliY + 28
   Call refreshScreen
  End If
  
  'go through all structures to see if there are any structures left in any of them
  For i = LBound(structureheight) + 1 To UBound(structureheight)
   If (structureheight(i) > 0) Then structuresleft = True
  Next i
  
  'if there are no structures left, start the next round by creating new buildings
  'and painting them. As well as moving the helicopter's co-ordinates back to the
  'top left. Because it's the new round, we increase the speed to make it harder.
  If structuresleft = False Then
   heliX = -48
   heliY = 0
   BitBlt myBackBuffer, 0, 0, 538, 368, 0, 0, 0, vbWhiteness
   Call CreateRandomBuilding
   Call paintBuilding
   speed = speed + 1
  End If
  
  'if the helicopter has collided with a building, we stop the helicopter from
  'moving by disabling this timer. We show a "Game Over" message box,check if the
  'best score is less than current score, if it is, store it to the file score.log,
  'unload this form and reshow the starting form.
  If heliCollision = True Then
   tmrHeli.Enabled = False
   MsgBox "Game Over", vbOKOnly
   
   Open frmStart.filename For Random Access Read Write As #1 Len = frmStart.recordlength
   Get #1, 1, BlankPlayer
   If BlankPlayer.score < points Then
    BlankPlayer.name = InputBox("Enter your name:", "Best Score")
    BlankPlayer.score = points
    Put #1, 1, BlankPlayer
   End If
   Close #1
   
   Unload Me
   frmStart.Show
  End If
End Sub

Private Function heliCollision() As Boolean 'this function is used in the timer above to see if the helicopter has collided with a building
 'for loop goes from 1 to 17 (each building)
 For i = LBound(structureheight) + 1 To UBound(structureheight)
 
  'this if statement determines if the front of the helicopter has made contact
  'with a building. If it's true, return true, remove the helicopter from the form,
  'place an explosion on the form,blit the backbuffer to the form, and leave the
  'function. If we never get in the if statement the function returns false
  If (heliY >= (frmGame.ScaleHeight - (structureheight(i) * 28))) And (heliX + 48 >= i * 28) Then
   heliCollision = True
   BitBlt myBackBuffer, heliX - speed, heliY, 48, 28, 0, 0, 0, vbBlackness
   BitBlt myBackBuffer, heliX + 20, heliY, 28, 28, myExplosion2Sprite, 0, 0, vbSrcPaint
   BitBlt Me.hdc, 0, 0, 538, 368, myBackBuffer, 0, 0, vbSrcCopy
   Exit For
  End If
  
 Next i
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 'if no missile is still in the air then perform shoot which uses the key we have
 'pressed as a parameter
 If freetoshoot Then shoot (KeyCode)
End Sub

Private Sub shoot(key As Integer) 'this procedure shoots a missile down or to the right
 'once we shoot we do not want to be able to shoot until the missile has hit a
 'building or gone off the screen, therefore false
 freetoshoot = False
 
 'key is the key which we have pressed
 Select Case key
  'when we press right this will be done
  Case vbKeyRight
  
   'start the missile in front of the helicopter because we are shooting right
   missileX = heliX + 48
   missileY = heliY + 14
   
    'repeat this do loop until the missile has hit the building or gone off the screen
    Do
      'wait 4 milli seconds, this is a procedure I made in modDelay
      delay (4)
      
      'clear the previous place the missile was in
      BitBlt myBackBuffer, missileX - 1, missileY, 3, 3, 0, 0, 0, vbBlackness

      'blit the new place where the missile sprite will be to the back-buffer
      BitBlt myBackBuffer, missileX, missileY, 3, 3, myMissileSprite, 0, 0, vbSrcPaint

      'now blit the backbuffer to the form...
      BitBlt Me.hdc, 0, 0, 532, 364, myBackBuffer, 0, 0, vbSrcCopy

      'increment the X co-ordinate to make it go the right next time we draw it
      missileX = missileX + 1
      
    Loop Until missileCollision


  'when we press down this will be done
  Case vbKeyDown
  
   'start the missile at the bottom of the helicopter because we are shooting down
   missileX = heliX + 24
   missileY = heliY + 28

    'repeat this do loop until the missile has hit the building or gone off the screen
    Do
     'wait 6 milli seconds, this is a procedure I made in modDelay
     delay (6)
     
     'clear the previous place the missile was in
     BitBlt myBackBuffer, missileX, missileY - 1, 3, 3, 0, 0, 0, vbBlackness

     'blit the new place where the missile sprite will be to the back-buffer
     BitBlt myBackBuffer, missileX, missileY, 3, 3, myMissileSprite, 0, 0, vbSrcPaint

     'now blit the backbuffer to the form...
     BitBlt Me.hdc, 0, 0, 538, 368, myBackBuffer, 0, 0, vbSrcCopy

    'increment the Y co-ordinate to make the missile go down next time we draw it
     missileY = missileY + 1

   Loop Until missileCollision
   
 End Select
 
 'the missile has gone off the screen or collided with a building,so now we can
 'shoot again
 freetoshoot = True
End Sub

Private Function missileCollision() As Boolean 'this function returns whether the missile has gone off the screen or collided with a building
 'if the missile has gone off the screen return true and skip all the stuff we need
 'to do when we hit a building
 If missileX > frmGame.ScaleWidth Or missileY > frmGame.ScaleHeight Then
  missileCollision = True
  GoTo outofscreen
 End If
 
 'go through all the buildings and check if the missile is in contact with a
 'building, if so, take one away from it's height, increment points by one, update
 'the form caption, return true, draw an explosion, and refresh the screen so that
 'the explosion does not appear anymore
 For i = LBound(structureheight) + 1 To UBound(structureheight)
  If (missileX + 2 >= i * 28) And (missileX - 2 <= (i + 1) * 28) And (missileY > (frmGame.ScaleHeight - (structureheight(i) * 28))) Then
   structureheight(i) = structureheight(i) - 1
   points = points + 1
   frmGame.Caption = "--Destroyer--  Points: " + Str(points)
   missileCollision = True
   Call explosion
   Call refreshScreen
  End If
 Next i
outofscreen:
End Function

Private Sub refreshScreen()
 'this function clears the screen from the helicopter down, and redraws everything
 'we need to see again.(helicopter and building)
 BitBlt myBackBuffer, 0, heliY - 28, frmGame.ScaleWidth, frmGame.ScaleHeight - (heliY - 28), 0, 0, 0, vbWhiteness
 BitBlt myBackBuffer, heliX, heliY, 48, 28, myHeliSprite, 0, 0, vbSrcPaint
 Call paintBuilding
End Sub

Private Sub explosion()
 'this function draws an initial explosion, waits 40 milliseconds, then draws
 'another one, and waits again, just to give it some kind of explosive effect
 BitBlt myBackBuffer, missileX - 14, missileY, 28, 28, myExplosion1Sprite, 0, 0, vbSrcPaint
 BitBlt Me.hdc, 0, 0, 538, 368, myBackBuffer, 0, 0, vbSrcCopy
 delay (40)
 BitBlt myBackBuffer, missileX - 14, missileY, 28, 28, 0, 0, 0, vbBlackness
 BitBlt myBackBuffer, missileX - 14, missileY, 28, 28, myExplosion2Sprite, 0, 0, vbSrcPaint
 BitBlt Me.hdc, 0, 0, 538, 368, myBackBuffer, 0, 0, vbSrcCopy
 delay (40)
End Sub

Private Sub Form_Unload(Cancel As Integer)
 'this clears up the memory we used to hold
 'the graphics and the buffers we made

 'Delete the bitmap surface that was in the backbuffer
 DeleteObject myBufferBMP

 'Delete the backbuffer HDC
 DeleteDC myBackBuffer

 'Delete the Sprite/Graphic HDC
 DeleteDC myHeliSprite
 DeleteDC myMissileSprite
 DeleteDC myBuildingSprite
 DeleteDC myExplosion2Sprite
 DeleteDC myExplosion1Sprite
End Sub
