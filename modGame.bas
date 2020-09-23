Attribute VB_Name = "modGame"
Option Explicit

'API Calls
'\/ Used to draw characters, map, etc
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'\/ Used to check for pressed keys
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'\/ Used to check pixels on the map's mask (check for collision)
Public Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
'\/ Used to control midi music
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
'\/ Used in function to get shortname of midi for playing the midi
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
'*****************

'Player Height and Width measurements for sprites
Public Const Play_W = 32 'Width
Public Const Play_H = 32 'Height
'Player Directions
Public Const P_Right = Play_W * 0 'Facing right
Public Const P_Left = Play_W * 1 'Facing Left
'Player poses
Public Const P_Stand = Play_H * 0 'Standing
Public Const P_Walk = Play_H * 1 'Walking
Public Const P_Punch = Play_H * 2 'Punching
Public Const P_Kick = Play_H * 3 'Kicking
Public Const P_Toss = Play_H * 4 'Tossing a gernade
Public Const P_Blast = Play_H * 5 'Firing weapon
Public Const P_Up = Play_H * 6 'Flying upwards
Public Const P_Fall = Play_H * 7 'Falling downwards
Public Const P_Fallen = Play_H * 8 'Laying on ground
Public Const P_GetUp = Play_H * 9 'Getting up from ground
Public Const P_OnFire = Play_H * 10 'Waving arms in random directions while on fire
'****************

'Weapon Types
Public Const WP_Blast = 0 'Energy ball (ClassicBob)
Public Const WP_Bomb = 1 'Gernade (different for each character)
Public Const WP_Poo = 2 'Poo particle (from PooGun or PoopSmith gernade)
Public Const WP_Flame = 3 'Flame spot (from flamethrower or FlamingBob gernade)
'****************

Public Type Player 'Player Typeset
Act As Long 'Whether player exists
Name As String 'Name of player
X As Double 'X axis position of player
Y As Double 'Y axis position of player
XS As Double 'X axis speed of player
YS As Double 'Y axis speed of player
Dir As Long 'Direction player faces
Ani As Long 'Player's pose
AniL As Long 'How long a player holds a pose (in cycles)
OnG As Long 'Whether or not player is on the ground
ID As Long 'ID number, same as index number
Jp As Long 'Whether or not player is in the air (used in keeping player in jumping poses)
HP As Double 'Hit points remaining
oldHP As Double
MP As Double 'Magic points remaing
tHP As Double 'Maximum hit points for player
Reload As Long 'Reload counter, uses to moderate gunfire
Dead As Long 'Whether player is dead or not (used for the jump and fall off of the screen seqeunce)
T As Long 'Character the player plays
OnFire As Long 'Whether or not player is on fire
OnFireL As Long 'How long the player will remain on fire (in cycles)
Moves(1 To 8) As Long 'Customizable keyset for the character (each number in array represents differennt key for a different function)

'AI Variables
MMood As Long
WMood As Long
AI As Boolean
End Type

Type Shot 'Typeset for shots (energy balls, poo particles, etc)
X As Double 'X axis position of shot
Y As Double 'Y azis position of shot
XS As Double 'X axis speed of shot
YS As Double 'Y axis speed of shot
ID As Long 'ID number of player to whom which the shot belongs
T As Long 'Type of shot
Bounce As Long 'How many bounces allowed (color in the case of the flame spot)
Fr As Long 'Current animation frame
FrL As Long 'Maximum animation frames
Act As Long 'Whether or not shot exists on the board
End Type

Type Explo 'Typeset for explosions
X As Long 'X axis position of explosion
Y As Long 'Y axis position of explosion
T As Long 'Type of explosion
Fr As Long 'Current frame of animation
FrL As Long 'Maximum animation frames
Act As Long 'Whether or not explosion exists on the board
End Type

Public P(1 To 2) As Player 'Player array (2 possible players)
Public S(1 To 100) As Shot 'Shot array (100 possible shots)
Public E(1 To 20) As Explo 'Explosion array (20 possible explosions)
Public Running As Boolean 'Whether or not game is running (used to stop mainloop function)
Public Paused As Boolean 'Whether or not game is paused (used to skip functions in mainloop while game is paused)
Public ShowMarker As Boolean 'Whether or not to show player color markers
Public LetChange As Boolean 'Whether or not to allow a player to change his/her character
Public dSound As New DS_Engine 'DirectSound engine (for sound obviously)

Function CheckWinner() As Long 'To check for the winner of the match
CheckWinner = 0 'Setting the value to no winner (for now)
If P(1).Act = False Then CheckWinner = 2 'If player 1 is dead then the winner is player 2
If P(2).Act = False Then CheckWinner = 1 'If player 2 is dead then the winner is player 1
If P(1).Act = False And P(2).Act = False Then CheckWinner = 3 'Both are dead, the match was a tie
End Function

Function ClearMem() 'Resets the arrays and maps back to normal
Dim I As Long 'Declare a counter variable
For I = 1 To 100 'Scroll through the shots
    S(I).Act = False 'Make sure the shot doesnt exist anymore
Next I
For I = 1 To 20 'Scroll through the explosions
    E(I).Act = False 'Make sure the explosions doesnt exist anymore
Next I
frmGame.bMask.Cls 'Clear the map's mask
frmGame.bSprite.Cls 'Clear the map's sprite
End Function

Function MoveShots()
'/\ The functions that moves the shots, checks for collisions with shots,
'and draws the shots, probably the biggest function in the game
'\/ Declaring the variable, here ill tell you what each does
'Counter, Counter, Counter, Opponent's ID, Numbers used to determine the direction
'of a shot, the last 4, used for checking collision (Square 2 Square)
Dim I As Long, I2 As Long, I3 As Long, opID As Long, _
XDr As Long, YDr As Long, X1, X2, Y1, Y2
For I = 1 To 100 'Scrolling through the shots (First level)

    If S(I).Act = True Then 'Only bother with this code \/ if the shot exists
        
        'Universal shot functions
        S(I).X = S(I).X + S(I).XS 'Move the shot on the X axis
        S(I).Y = S(I).Y + S(I).YS 'Move the shot on the Y axis
        
        If S(I).X > frmGame.bMask.ScaleWidth Then 'If shot goes off of the right map boundry
            If S(I).T = WP_Bomb Then 'If the shot is a gernade then
                S(I).XS = -S(I).XS   'bounce it back
            Else
                S(I).Act = False 'If it isnt, destroy it
            End If
        End If
        If S(I).X < -32 Then 'If shot goes off of the left map boundry
            If S(I).T = WP_Bomb Then 'If the shot is a gernade then
                S(I).XS = -S(I).XS   'bounce it back
            Else
                S(I).Act = False 'If it isnt, destroy it
            End If
        End If
        
        If S(I).Y > frmGame.bMask.ScaleHeight Then S(I).Act = False 'if shot has
        'fallen off of the map, destory it (if it hit the bottom boundry)
        
        opID = IIf(S(I).ID = 1, 2, 1) 'get the id opposite from the shot's id
        'this will be the only id (player) that will recieve damage if hit
        '*********** end of universal shot functions
        
        'Special shot functions
        Select Case S(I).T
        Case WP_Blast 'The special functions for the energy ball
            
            For I2 = 1 To 2
                'Scroll through the player array and find out if the shot hit
                'either character
                X1 = S(I).X 'Setting the position for the shot
                Y1 = S(I).Y
                X2 = P(I2).X + 10 'Setting the position for the player
                Y2 = P(I2).Y
                If X2 > X1 And X2 < X1 + 16 And Y2 > Y1 And Y2 < Y1 + 8 Or X1 > X2 And X1 < X2 + 12 And Y1 > Y2 And Y1 < Y2 + Play_H Then
                    '/\ check for collision
                    P(I2).HP = P(I2).HP - 100 'Deduct from the hit player's hit points
                    S(I).Act = False 'Destroy the shot
                End If
            Next I2
            
            If GetPixel(frmGame.bMask.hDC, S(I).X, S(I).Y) = vbBlack Then
                'If the energy ball hit terrain then destroy the shot and
                'deform the terrain
                S(I).Act = False
                frmGame.bMask.DrawWidth = 10 'Set the whole size
                frmGame.bSprite.DrawWidth = 13
                frmGame.bMask.PSet (S(I).X, S(I).Y), vbWhite 'Modify the map mask
                frmGame.bSprite.PSet (S(I).X, S(I).Y), vbBlack 'Modify the map sprite
            End If
            
            '\/ Draws the energy ball (shot)
            BitBlt frmGame.Board.hDC, S(I).X, S(I).Y, 16, 8, frmGame.FM.hDC, 0, IIf(S(I).XS < 0, 8, 0), vbSrcAnd
            BitBlt frmGame.Board.hDC, S(I).X, S(I).Y, 16, 8, frmGame.FS.hDC, 0, IIf(S(I).XS < 0, 8, 0), vbSrcInvert
        
        Case WP_Bomb 'Special functions for gernades
            S(I).YS = S(I).YS + 0.1 'Apply gravity to the gernade
            
            S(I).Fr = S(I).Fr + 1: If S(I).Fr > S(I).FrL Then S(I).Fr = 0
            '/\ Move the animation for the gernade (rolling)
            
            'Draw the gernade \/
            BitBlt frmGame.Board.hDC, S(I).X, S(I).Y, 8, 8, frmGame.GM.hDC, S(I).Fr * 8, 0, vbSrcAnd
            BitBlt frmGame.Board.hDC, S(I).X, S(I).Y, 8, 8, frmGame.GS.hDC, S(I).Fr * 8, 0, vbSrcInvert
            
            For I2 = 1 To 2 'Scroll through the players to see if they were hit
                X1 = S(I).X
                X2 = P(I2).X + 10
                Y1 = S(I).Y
                Y2 = P(I2).Y
                If X2 > X1 And X2 < X1 + 16 And Y2 > Y1 And Y2 < Y1 + 8 Or X1 > X2 And X1 < X2 + 12 And Y1 > Y2 And Y1 < Y2 + Play_H Then
                    dSound.PlaySound 2, False
                    DoExplo S(I).X - 12, S(I).Y - 12, 0, 8
                    P(I2).HP = P(I2).HP - 50
                    If P(S(I).ID).T = 0 Then P(I2).HP = P(I2).HP - 100
                    If P(S(I).ID).T <> 0 Then
                        For I3 = 1 To 5
                            pBlast P(S(I).ID), S(I).X, S(I).Y + 6, Int(Rnd * 3) + 0.25 * IIf(Int(Rnd * 2) = 0, -1, 1), -2.5
                        Next I3
                        For I3 = 1 To 5
                            pBlast P(S(I).ID), S(I).X, S(I).Y - 6, Int(Rnd * 3) + 0.25 * IIf(Int(Rnd * 2) = 0, -1, 1), -2.5
                        Next I3
                        For I3 = 1 To 5
                            pBlast P(S(I).ID), S(I).X - 6, S(I).Y, Int(Rnd * 3) + 0.25 * IIf(Int(Rnd * 2) = 0, -1, 1), -2.5
                        Next I3
                        For I3 = 1 To 5
                            pBlast P(S(I).ID), S(I).X + 6, S(I).Y, Int(Rnd * 3) + 0.25 * IIf(Int(Rnd * 2) = 0, -1, 1), -2.5
                        Next I3
                    Else
                        frmGame.bMask.DrawWidth = 22
                        frmGame.bSprite.DrawWidth = 25
                        frmGame.bMask.PSet (S(I).X, S(I).Y), vbWhite
                        frmGame.bSprite.PSet (S(I).X, S(I).Y), vbBlack
                    End If
                    S(I).Act = False
                End If
            Next I2
            If S(I).Bounce > 0 Then
                If GetPixel(frmGame.bMask.hDC, S(I).X + S(I).XS, S(I).Y + S(I).YS) = vbBlack Then
                    XDr = 0
                    YDr = 0
                    If GetPixel(frmGame.bMask.hDC, S(I).X + S(I).XS, S(I).Y) Then XDr = 1
                    If GetPixel(frmGame.bMask.hDC, S(I).X, S(I).Y + S(I).YS) Then YDr = 1
                    If XDr = 1 Then S(I).XS = S(I).XS * -0.8
                    If YDr = 1 Then S(I).YS = S(I).YS * -0.8
                    S(I).Bounce = S(I).Bounce - 1
                End If
            Else
                If GetPixel(frmGame.bMask.hDC, S(I).X, S(I).Y) = vbBlack Then
                    dSound.PlaySound 2, False
                    DoExplo S(I).X - 12, S(I).Y - 12, 0, 8
                    If P(S(I).ID).T <> 0 Then
                        For I3 = 1 To 5
                            pBlast P(S(I).ID), S(I).X, S(I).Y + 6, Int(Rnd * 3) + 0.25 * IIf(Int(Rnd * 2) = 0, -1, 1), -2.5
                        Next I3
                        For I3 = 1 To 5
                            pBlast P(S(I).ID), S(I).X, S(I).Y - 6, Int(Rnd * 3) + 0.25 * IIf(Int(Rnd * 2) = 0, -1, 1), -2.5
                        Next I3
                        For I3 = 1 To 5
                            pBlast P(S(I).ID), S(I).X - 6, S(I).Y, Int(Rnd * 3) + 0.25 * IIf(Int(Rnd * 2) = 0, -1, 1), -2.5
                        Next I3
                        For I3 = 1 To 5
                            pBlast P(S(I).ID), S(I).X + 6, S(I).Y, Int(Rnd * 3) + 0.25 * IIf(Int(Rnd * 2) = 0, -1, 1), -2.5
                        Next I3
                    Else
                        frmGame.bMask.DrawWidth = 22
                        frmGame.bSprite.DrawWidth = 25
                        frmGame.bMask.PSet (S(I).X, S(I).Y), vbWhite
                        frmGame.bSprite.PSet (S(I).X, S(I).Y), vbBlack
                    End If
                    S(I).Act = False
                End If
            End If
        Case WP_Poo
            S(I).YS = S(I).YS + 0.35
            frmGame.Board.DrawWidth = 5
            frmGame.Board.PSet (S(I).X, S(I).Y), &H4080&
            For I2 = 1 To 2
                X1 = S(I).X - 2.5
                X2 = P(I2).X + 10
                Y1 = S(I).Y - 2.5
                Y2 = P(I2).Y
                If X2 > X1 And X2 < X1 + 5 And Y2 > Y1 And Y2 < Y1 + 5 Or X1 > X2 And X1 < X2 + 12 And Y1 > Y2 And Y1 < Y2 + Play_H Then
                    P(I2).HP = P(I2).HP - 15
                    S(I).Act = False
                End If
            Next I2
            If S(I).Bounce > 0 Then
                If GetPixel(frmGame.bMask.hDC, S(I).X + S(I).XS, S(I).Y + S(I).YS) = vbBlack Then
                    XDr = 0
                    YDr = 0
                    If GetPixel(frmGame.bMask.hDC, S(I).X + S(I).XS, S(I).Y) Then XDr = 1
                    If GetPixel(frmGame.bMask.hDC, S(I).X, S(I).Y + S(I).YS) Then YDr = 1
                    If XDr = 1 Then S(I).XS = S(I).XS * -0.8
                    If YDr = 1 Then S(I).YS = S(I).YS * -0.8
                    S(I).Bounce = S(I).Bounce - 1
                End If
            Else
                If GetPixel(frmGame.bMask.hDC, S(I).X, S(I).Y) = vbBlack Then
                    S(I).Act = False
                    frmGame.bMask.DrawWidth = 5
                    frmGame.bSprite.DrawWidth = 5
                    frmGame.bMask.PSet (S(I).X - S(I).XS / 2, S(I).Y - S(I).YS / 2), vbBlack
                    frmGame.bSprite.PSet (S(I).X - S(I).XS / 2, S(I).Y - S(I).YS / 2), &H4080&
                End If
            End If
        Case WP_Flame
            S(I).YS = S(I).YS - 0.05
            S(I).Fr = S(I).Fr + 1: If S(I).Fr > S(I).FrL Then S(I).Act = False
            frmGame.Board.DrawWidth = 5
            frmGame.Board.PSet (S(I).X, S(I).Y), S(I).Bounce
            For I2 = 1 To 2
                If P(I2).ID <> S(I).ID Then
                    X1 = S(I).X - 2.5
                    X2 = P(I2).X + 10
                    Y1 = S(I).Y - 2.5
                    Y2 = P(I2).Y
                    If X2 > X1 And X2 < X1 + 5 And Y2 > Y1 And Y2 < Y1 + 5 Or X1 > X2 And X1 < X2 + 12 And Y1 > Y2 And Y1 < Y2 + Play_H Then
                        If S(I).ID = 44 Then Exit For
                        P(I2).HP = P(I2).HP - 1
                        S(I).Fr = S(I).Fr + 5
                        P(I2).OnFire = True
                        P(I2).OnFireL = 50
                    End If
                End If
            Next I2
        End Select
    End If
Next I
End Function

Function pBomb(P As Player)
Dim I As Long
If P.OnFire = True Then Exit Function
If P.Reload > 0 Then Exit Function
If P.MP < 75 Then Exit Function
P.Reload = 20
P.Ani = P_Toss
P.AniL = 15
For I = 1 To 100
    If S(I).Act = False Then
        S(I).X = IIf(P.Dir = P_Left, P.X - 6, P.X + 36)
        S(I).Y = P.Y - 9
        S(I).XS = IIf(P.Dir = P_Left, -2.5, 2.5)
        S(I).YS = -2.5
        S(I).T = WP_Bomb
        S(I).Act = True
        S(I).Bounce = 2
        S(I).ID = P.ID
        S(I).Fr = 0
        S(I).FrL = 3
        P.MP = P.MP - 75
        Exit For
    End If
Next I
End Function

Function pBlast(P As Player, Optional X As Double = -1, Optional Y As Double = -1, Optional XS As Double = 25.5, Optional YS As Double = 25.5)
Dim I As Long
If P.OnFire = True Then Exit Function
If P.Reload > 0 Then Exit Function
If P.MP < 5 Then Exit Function
P.Ani = P_Blast
P.AniL = 5
Select Case P.T
Case 0 'Classic
    For I = 1 To 100
        If S(I).Act = False Then
            S(I).X = IIf(P.Dir = P_Left, P.X - 8, P.X + 24)
            S(I).Y = P.Y + 8
            S(I).XS = IIf(P.Dir = P_Left, -3.5, 3.5)
            S(I).YS = 0
            S(I).T = WP_Blast
            S(I).Act = True
            P.Reload = 18
            P.MP = P.MP - 20
            dSound.PlaySound 6, False
            Exit For
        End If
    Next I
Case 1 'Poopsmith
    For I = 1 To 100
        If S(I).Act = False Then
            S(I).X = IIf(X = -1, IIf(P.Dir = P_Left, P.X, P.X + 27), X)
            S(I).Y = IIf(Y = -1, P.Y - 3, Y)
            S(I).XS = IIf(XS = 25.5, IIf(P.Dir = P_Left, -(Rnd * 2.5 + 1), (Rnd * 2.5 + 1)), XS)
            S(I).YS = IIf(YS = 25.5, -(Rnd * 2.5 + 1), YS)
            S(I).T = WP_Poo
            S(I).ID = P.ID
            S(I).Act = True
            S(I).Bounce = 2
            P.Reload = IIf(X = -1, 6, 0)
            P.MP = P.MP - IIf(X = -1, 5, 0)
            If X = -1 Then dSound.PlaySound 5, False
            Exit For
        End If
    Next I
Case 2 'Flaming Bob
    For I = 1 To 100
        If S(I).Act = False Then
            S(I).X = IIf(X = -1, IIf(P.Dir = P_Left, P.X + 2, P.X + 25), X)
            S(I).Y = IIf(Y = -1, P.Y + 12, Y)
            S(I).XS = IIf(XS = 25.5, IIf(P.Dir = P_Left, -(Rnd * 2.5 + 1), (Rnd * 2.5 + 1)), XS)
            S(I).YS = 0
            S(I).T = WP_Flame
            S(I).Act = True
            S(I).Fr = 0
            S(I).ID = P.ID
            S(I).FrL = Int(Rnd * 10) + 25
            S(I).Bounce = IIf(Int(Rnd * 2) = 0, &HFF&, IIf(Int(Rnd * 2) = 0, &H80FF&, &HFFFF&))
            P.Reload = IIf(X = -1, 2, 0)
            P.MP = P.MP - IIf(X = -1, 1, 0)
            Exit For
        End If
    Next I
End Select
End Function

Function pPunch(Pl As Player)
Dim opID As Long
If Pl.OnFire = True Then Exit Function
If Pl.Reload > 0 Then Exit Function
If Pl.MP < 2 Then Exit Function
opID = IIf(Pl.ID = 1, 2, 1)
Pl.Ani = P_Punch
Pl.AniL = 10
Pl.Reload = 10
Dim X1, X2, Y1, Y2
X1 = Pl.X
X2 = P(opID).X
Y1 = Pl.Y
Y2 = P(opID).Y
If X2 > X1 And X2 < X1 + Play_W And Y2 > Y1 And Y2 < Y1 + Play_H Or X1 > X2 And X1 < X2 + Play_W And Y1 > Y2 And Y1 < Y2 + Play_H Then
    P(opID).HP = P(opID).HP - 15
    P(opID).Ani = P_Fallen
    P(opID).AniL = 10
End If
End Function

Function pKick(Pl As Player)
Dim opID As Long
If Pl.OnFire = True Then Exit Function
If Pl.Reload > 0 Then Exit Function
If Pl.MP < 5 Then Exit Function
opID = IIf(Pl.ID = 1, 2, 1)
Pl.Ani = P_Kick
Pl.Reload = 15
Pl.AniL = 15
Dim X1, X2, Y1, Y2
X1 = Pl.X
X2 = P(opID).X
Y1 = Pl.Y
Y2 = P(opID).Y
If X2 > X1 And X2 < X1 + Play_W And Y2 > Y1 And Y2 < Y1 + Play_H Or X1 > X2 And X1 < X2 + Play_W And Y1 > Y2 And Y1 < Y2 + Play_H Then
    P(opID).HP = P(opID).HP - 25
    P(opID).Ani = P_Fallen
    P(opID).AniL = 15
End If
End Function

Function DoKeys(P As Player, Data)
If P.OnFire = True Then Exit Function
Dim I As Long
For I = 1 To Len(Data)
    Select Case Mid(Data, I, 1)
    Case "J": pJump P
    Case "L": pWalk P, P_Left
    Case "R": pWalk P, P_Right
    Case "F": pPunch P
    Case "K": pKick P
    Case "B": pBlast P
    Case "G": pBomb P
    Case "C"
        If P.Reload > 0 Then Exit Function
        P.Reload = 50
        P.T = P.T + 1: If P.T > 2 Then P.T = 0
        Data = Data & "C"
    End Select
Next I
End Function

Function GetKeys(P As Player, J, L, R, Pn, K, B, G, C)
Dim Data As String
If P.AI = True Then
    DoAI P
    Exit Function
End If
If P.OnFire = True Then Exit Function
If GetAsyncKeyState(J) Then
    pJump P
    Data = Data & "J"
End If
If GetAsyncKeyState(L) Then
    pWalk P, P_Left
    Data = Data & "L"
End If
If GetAsyncKeyState(R) Then
    pWalk P, P_Right
    Data = Data & "R"
End If
If GetAsyncKeyState(Pn) Then
    pPunch P
    Data = Data & "F"
End If
If GetAsyncKeyState(K) Then
    pKick P
    Data = Data & "K"
End If
If GetAsyncKeyState(B) Then
    pBlast P
    Data = Data & "B"
End If
If GetAsyncKeyState(G) Then
    pBomb P
    Data = Data & "G"
End If
If GetAsyncKeyState(C) And LetChange = True Then
    If P.Reload > 0 Then Exit Function
    P.Reload = 50
    P.T = P.T + 1: If P.T > 2 Then P.T = 0
    Data = Data & "C"
End If
If Data = "" Then Exit Function
Data = "PLY|¿|" & P.X & "," & P.Y & "|¿|" & Data
If frmGame.wsClient.State = sckConnected Then frmGame.wsClient.SendData Data
If frmGame.wsHost.State = sckConnected Then frmGame.wsHost.SendData Data
End Function

Function pWalk(P As Player, D As Long)
P.OnG = False
P.XS = IIf(D = P_Right, 2, -2)
P.Dir = D
If P.Jp = True Then Exit Function
P.Ani = IIf(P.Ani = P_Stand, P_Walk, P_Stand)
P.AniL = 2
End Function

Function pJump(P As Player)
If P.OnFire = True Then Exit Function
If P.Act = True And P.OnG = True Then
    P.Jp = True
    P.OnG = False
    P.Ani = P_Up
    P.YS = -3.5
End If
End Function

Function NewPlayer(P As Player, X, Y, ID, T, HP, Name As String, AI)
P.Reload = 0
P.X = X
P.Y = Y
P.XS = 0
P.YS = 0
P.Ani = P_Stand
P.AniL = 0
P.Dir = P_Right
P.OnG = False
P.Name = Name
P.ID = ID
P.HP = HP
P.oldHP = HP
P.tHP = HP
P.T = T
P.Act = True
P.Dead = False
P.AI = IIf(AI = 0, False, True)
End Function

Function MovePlayers()
Dim I As Long, I2 As Long, Sle As Long
For I = 1 To 2
    If P(I).Act = True Then
        If P(I).Dead = False And P(I).HP <= 0 And P(I).OnFire = False Then
            dSound.PlaySound 4, False
            P(I).Dead = True
            pJump P(I)
        End If
        P(I).OnFireL = P(I).OnFireL - 1: If P(I).OnFireL <= 0 Then P(I).OnFireL = 0
        If P(I).OnFireL <= 0 Then P(I).OnFire = False
        If P(I).OnFire = True Then
            If P(I).OnFireL > 10 And P(I).OnFireL < 20 Or P(I).OnFire < 40 And P(I).OnFire > 30 Then
                P(I).Dir = P_Left
                pWalk P(I), P_Left
                P(I).Ani = P_OnFire
            ElseIf P(I).OnFireL < 10 Or P(I).OnFire > 20 And P(I).OnFire < 30 Or P(I).OnFire < 50 And P(I).OnFire > 40 Then
                P(I).Dir = P_Right
                pWalk P(I), P_Right
                P(I).Ani = P_OnFire
            End If
            Dim Dum As Player
            Dum.Act = True
            Dum.T = 2
            Dum.ID = 44
            For I2 = 1 To 6
                Dum.MP = 80
                Dum.Reload = 0
                pBlast Dum, P(I).X + 16, P(I).Y + 16, Int(Rnd * 2) * IIf(Int(Rnd * 2) = 0, -1, 1), -0.5
            Next I2
            P(I).Ani = P_OnFire
            P(I).AniL = 5
        End If
        P(I).Reload = P(I).Reload - 1
        If P(I).Reload < 0 Then P(I).Reload = 0
        P(I).MP = P(I).MP + 0.2: If P(I).MP > 80 Then P(I).MP = 80
        P(I).AniL = P(I).AniL - 1
        If P(I).AniL <= 0 Then
            Select Case P(I).Ani
            Case P_OnFire And P(I).OnFire = False
                P(I).Ani = P_Fallen
                P(I).AniL = 25
            Case P_Fallen
                P(I).Ani = P_GetUp
                P(I).AniL = 20
            Case Else
                P(I).Ani = P_Stand
                P(I).AniL = 5
            End Select
        End If
        If P(I).Jp = True Then P(I).Ani = P_Up
        If P(I).YS < -0.5 Then P(I).Ani = P_Up
        If P(I).YS > 2 Then P(I).Ani = P_Fall
        If P(I).Dead = False And P(I).OnG = False And GetPixel(frmGame.bMask.hDC, P(I).X + Play_W / 2, P(I).Y) = vbBlack Then
            P(I).Y = P(I).Y + 1
            P(I).YS = -P(I).YS
        End If
        If P(I).Dead = False And P(I).OnG = False And GetPixel(frmGame.bMask.hDC, P(I).X + Play_W / 2, P(I).Y + Play_H) = vbBlack Then
            P(I).Jp = False
            P(I).Y = P(I).Y - 1
            P(I).YS = 0
            P(I).OnG = True
        End If
        If GetPixel(frmGame.bMask.hDC, P(I).X + P(I).XS + Play_W * 0.8, P(I).Y + Play_H / 2) = vbBlack Then
            P(I).XS = 0
        End If
        If GetPixel(frmGame.bMask.hDC, P(I).X + P(I).XS + Play_W * 0.2, P(I).Y + Play_H / 2) = vbBlack Then
            P(I).XS = 0
        End If
        If P(I).X + P(I).XS + (Play_W * 0.75) > 320 Then P(I).XS = 0
        If P(I).X + P(I).XS < -(Play_W / 4) Then P(I).XS = 0
        P(I).X = P(I).X + P(I).XS
        If P(I).OnG = False Then
            P(I).YS = P(I).YS + 0.1
        End If
        P(I).Y = P(I).Y + P(I).YS
        If P(I).OnG = True Then
            P(I).XS = P(I).XS * 0.75
        End If
        If P(I).Y > 240 Then P(I).Act = False
        If ShowMarker = True Then
            frmGame.Board.DrawWidth = 1
            frmGame.Board.Line (P(I).X, P(I).Y)-(P(I).X + 32, P(I).Y + 32), IIf(P(I).ID = 1, vbRed, vbBlue), B
        End If
        BitBlt frmGame.Board.hDC, P(I).X, P(I).Y, Play_W, Play_H, frmGame.SM(P(I).T).hDC, P(I).Dir, P(I).Ani, vbSrcAnd
        BitBlt frmGame.Board.hDC, P(I).X, P(I).Y, Play_W, Play_H, frmGame.SS(P(I).T).hDC, P(I).Dir, P(I).Ani, vbSrcInvert
        frmGame.pHP(I - 1).Cls
        frmGame.pHP(I - 1).Line (0, 0)-(frmGame.pHP(I - 1).ScaleWidth * (P(I).HP / P(I).tHP), frmGame.pHP(I - 1).Height), &H80FF80, BF
        frmGame.pEN(I - 1).Cls
        frmGame.pEN(I - 1).Line (0, 0)-(frmGame.pEN(I - 1).ScaleWidth * (P(I).MP / 80), frmGame.pEN(I - 1).Height), &HFFC0FF, BF
    End If
Next I
End Function

Function MoveExplo()
Dim I As Long
For I = 1 To 20
    If E(I).Act = True Then
        E(I).Fr = E(I).Fr + 1: If E(I).Fr > (E(I).FrL - 1) Then E(I).Act = False
        BitBlt frmGame.Board.hDC, E(I).X, E(I).Y, 32, 32, frmGame.EM(E(I).T).hDC, E(I).Fr * 32, 0, vbSrcAnd
        BitBlt frmGame.Board.hDC, E(I).X, E(I).Y, 32, 32, frmGame.ES(E(I).T).hDC, E(I).Fr * 32, 0, vbSrcInvert
    End If
Next I
End Function

Function DoExplo(X As Long, Y As Long, T As Long, FrL As Long)
Dim I As Long
For I = 1 To 20
    If E(I).Act = False Then
        E(I).FrL = FrL
        E(I).Fr = 0
        E(I).T = T
        E(I).X = X
        E(I).Y = Y
        E(I).Act = True
        Exit For
    End If
Next I
End Function

'Not Mine ********** Forget author, but thanks whoever wrote it
Public Function GetShortFileName(ByVal Filename As String) As String
    'converts a long file and path name to o
    '     ld DOS format
    'PARAMETERS
    'FileName = the path or filename to conv
    '     ert
    'RETURNS
    'String = the DOS compatible name for th
    '     at particular FileName
    'USES
    'KERNEL32 API call GetShortPathNameA
    'CONSIDERATIONS
    'short filename equivalents should only
    '     be used with non
    'Win95 programs
    Dim rc As Long
    Dim ShortPath As String
    Const PATH_LEN& = 164
    'get the short filename
    ShortPath = String$(PATH_LEN + 1, 0)
    rc = GetShortPathName(Filename, ShortPath, PATH_LEN)
    GetShortFileName = Left$(ShortPath, rc)
End Function
