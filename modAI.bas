Attribute VB_Name = "modAI"
Option Explicit

Function DoAI(Pl As Player)
Dim opp As Long, I As Long

If Int(Rnd * 5) < 2 Then Exit Function
opp = IIf(Pl.ID = 1, 2, 1)
If Int(Rnd * 50) < 5 Then Pl.WMood = IIf(Pl.WMood = 0, 1, 0)
If Int(Rnd * 50) < 5 Then Pl.MMood = 0
If P(opp).Y < Pl.Y - 16 And Int(Rnd * 5) = 0 Then pJump Pl

For I = 1 To 100
    If S(I).Act = True Then
        If S(I).X > Pl.X - 50 And S(I).X < Pl.X + 82 And S(I).Y > Pl.Y - 6 And S(I).Y < Pl.Y + 40 Then
            If Pl.OnG = True Then
                pJump Pl
            Else
                Select Case S(I).XS
                Case Is > 0
                    pWalk Pl, P_Right
                Case Is < 0
                    pWalk Pl, P_Left
                End Select
            End If
        Exit Function
        End If
    End If
Next I

If P(opp).HP < P(opp).oldHP Then
    P(opp).oldHP = P(opp).HP
    Pl.MMood = 1
End If

If P(opp).Y > Pl.Y - 16 And P(opp).Y < Pl.Y + 48 Then
    If P(opp).X > Pl.X - 16 And P(opp).X < Pl.X + 48 Then
        If Int(Rnd * 2) = 0 Then
            pPunch Pl
        Else
            pKick Pl
        End If
    End If
End If

If Pl.MMood = 0 Then
If P(opp).X + 16 < Pl.X + 16 And P(opp).X + 32 < Pl.X + 16 Then
    Pl.Dir = P_Left
    If Int(Rnd * 5) < 3 Then pWalk Pl, P_Left
    If GetPixel(frmGame.bMask.hDC, Pl.X - 16, Pl.Y - 8) = vbBlack Or GetPixel(frmGame.bMask.hDC, Pl.X - 16, Pl.Y + 48) = vbWhite Or GetPixel(frmGame.bMask.hDC, Pl.X - 16, Pl.Y + 16) = vbBlack Then pJump Pl
    Select Case Pl.WMood
    Case 1 'weapons
        If P(opp).X < Pl.X - 75 And P(opp).X > Pl.X - 115 And P(opp).Y > Pl.Y - 32 Then
            pBomb Pl
        End If
        If P(opp).Y > Pl.Y - 8 And P(opp).Y < Pl.Y + 48 Then
            pBlast Pl
        End If
    End Select
End If
If P(opp).X + 16 > Pl.X + 16 And P(opp).X + 32 > Pl.X + 16 Then
    Pl.Dir = P_Right
    If Int(Rnd * 5) < 3 Then pWalk Pl, P_Right
    If GetPixel(frmGame.bMask.hDC, Pl.X + 48, Pl.Y - 8) = vbBlack Or GetPixel(frmGame.bMask.hDC, Pl.X + 48, Pl.Y + 48) = vbWhite Or GetPixel(frmGame.bMask.hDC, Pl.X + 48, Pl.Y + 16) = vbBlack Then pJump Pl
    Select Case Pl.WMood
    Case 1 'weapons
        If P(opp).X < Pl.X + 109 And P(opp).X > Pl.X + 147 And P(opp).Y > Pl.Y - 32 Then
            pBomb Pl
        End If
        If P(opp).Y > Pl.Y - 8 And P(opp).Y < Pl.Y + 48 Then
            pBlast Pl
        End If
    End Select
End If
End If

If Pl.MMood = 1 Then
If P(opp).X + 16 < Pl.X + 16 And P(opp).X + 32 < Pl.X + 16 Then
    Pl.Dir = P_Right
    If Int(Rnd * 5) < 3 Then pWalk Pl, P_Right
    If GetPixel(frmGame.bMask.hDC, Pl.X - 16, Pl.Y - 8) = vbBlack Or GetPixel(frmGame.bMask.hDC, Pl.X - 16, Pl.Y + 48) = vbWhite Then pJump Pl
End If
If P(opp).X + 16 > Pl.X + 16 And P(opp).X + 32 > Pl.X + 16 Then
    Pl.Dir = P_Left
    If Int(Rnd * 5) < 3 Then pWalk Pl, P_Left
    If GetPixel(frmGame.bMask.hDC, Pl.X + 48, Pl.Y - 8) = vbBlack Or GetPixel(frmGame.bMask.hDC, Pl.X + 48, Pl.Y + 48) = vbWhite Then pJump Pl
End If
End If
End Function
