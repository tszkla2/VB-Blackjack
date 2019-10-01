'Tom Szklarzewski
'Final Project - Blackjack

Public Class Form1
    Dim intUserDraw1 As Integer
    Dim intUserDraw2 As Integer
    Dim intDealerDraw1 As Integer
    Dim intDealerDraw2 As Integer
    Dim intDealerDraw3 As Integer
    Dim intDealerDraw4 As Integer
    Dim intDealerDraw5 As Integer
    Dim intDealerDraw6 As Integer
    Dim intDealerDraw7 As Integer
    Dim intDealerDraw8 As Integer
    Dim intDealerDraw9 As Integer
    Dim intHitDraw3 As Integer
    Dim intHitDraw4 As Integer
    Dim intHitDraw5 As Integer
    Dim intHitDraw6 As Integer
    Dim intHitDraw7 As Integer
    Dim intHitDraw8 As Integer
    Dim intHitDraw9 As Integer
    Dim intCounter As Integer
    Dim strAce As String

    Private Sub btnShuffle_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnShuffle.Click
        Randomize()
        'Clear All Cards
        Me.lblDealerCard1.Text = Nothing
        Me.lblDealerCard2.Text = Nothing
        Me.lblDealerCard3.Text = Nothing
        Me.lblDealerCard4.Text = Nothing
        Me.lblDealerCard5.Text = Nothing
        Me.lblDealerCard6.Text = Nothing
        Me.lblDealerCard7.Text = Nothing
        Me.lblDealerCard8.Text = Nothing
        Me.lblDealerCard9.Text = Nothing
        Me.lblDealerScore.Text = Nothing
        Me.lblUserCard3.Text = Nothing
        Me.lblUserCard4.Text = Nothing
        Me.lblUserCard5.Text = Nothing
        Me.lblUserCard6.Text = Nothing
        Me.lblUserCard7.Text = Nothing
        Me.lblUserCard8.Text = Nothing
        Me.lblUserCard9.Text = Nothing
        Me.lblPlayerScore.Text = Nothing
        Me.lblMessage.Text = Nothing
        lblUserCard3.Visible = False
        lblUserCard4.Visible = False
        lblUserCard5.Visible = False
        lblUserCard6.Visible = False
        lblUserCard7.Visible = False
        lblUserCard8.Visible = False
        lblUserCard9.Visible = False
        lblDealerCard2.Visible = False
        lblDealerCard3.Visible = False
        lblDealerCard4.Visible = False
        lblDealerCard5.Visible = False
        lblDealerCard6.Visible = False
        lblDealerCard7.Visible = False
        lblDealerCard8.Visible = False
        lblDealerCard9.Visible = False
        intCounter = 0
        strAce = 0

        'Random Number
        intUserDraw1 = Int(Rnd() * 10) + 1
        intUserDraw2 = Int(Rnd() * 10) + 1
        intDealerDraw1 = Int(Rnd() * 10) + 1

        'Player Draw Cards
        Me.lblUserCard1.Text = intUserDraw1
        Me.lblUserCard2.Text = intUserDraw2
        Me.lblDealerCard1.Text = intDealerDraw1
        If intUserDraw1 = 1 Then
            Me.lblUserCard1.Text = "A"
        End If
        If intUserDraw2 = 1 Then
            Me.lblUserCard2.Text = "A"
        End If
        lblDealerCard1.Visible = True
        lblUserCard1.Visible = True
        lblUserCard2.Visible = True
        If intUserDraw1 = 1 Then
            strAce = Val(InputBox("Ace Value 1 or 11?"))
            Me.lblUserCard1.Text = strAce
            intUserDraw1 = strAce
        End If
        If intUserDraw2 = 1 Then
            strAce = Val(InputBox("Ace Value 1 or 11?"))
            Me.lblUserCard2.Text = strAce
            intUserDraw2 = strAce
        End If
        Me.lblPlayerScore.Text = (intUserDraw1 + intUserDraw2)
        Me.lblDealerScore.Text = intDealerDraw1
        If Me.lblPlayerScore.Text > 21 Then
            Me.lblMessage.Text = "Player Busts! Dealer Wins!"
        End If
    End Sub

    Private Sub btnStand_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnStand.Click
        'Clear Score
        Me.lblDealerScore.Text = Nothing

        'Random Number
        intDealerDraw2 = Int(Rnd() * 10) + 1
        intDealerDraw3 = Int(Rnd() * 10) + 1
        intDealerDraw4 = Int(Rnd() * 10) + 1
        intDealerDraw5 = Int(Rnd() * 10) + 1
        intDealerDraw6 = Int(Rnd() * 10) + 1
        intDealerDraw7 = Int(Rnd() * 10) + 1
        intDealerDraw8 = Int(Rnd() * 10) + 1
        intDealerDraw9 = Int(Rnd() * 10) + 1

        'Dealer Draw
        Me.lblDealerCard2.Text = intDealerDraw2
        lblDealerCard2.Visible = True
        Me.lblDealerScore.Text = Val(intDealerDraw1 + intDealerDraw2)
        If lblDealerScore.Text < 17 Then
            Me.lblDealerCard3.Text = intDealerDraw3
            Me.lblDealerCard3.Visible = True
            Me.lblDealerScore.Text = Val(intDealerDraw1 + intDealerDraw2 + intDealerDraw3)
        End If
        If lblDealerScore.Text < 17 Then
            Me.lblDealerCard4.Text = intDealerDraw4
            Me.lblDealerCard4.Visible = True
            Me.lblDealerScore.Text = Val(intDealerDraw1 + intDealerDraw2 + intDealerDraw3 + intDealerDraw4)
        End If
        If lblDealerScore.Text < 17 Then
            Me.lblDealerCard5.Text = intDealerDraw5
            Me.lblDealerCard5.Visible = True
            Me.lblDealerScore.Text = Val(intDealerDraw1 + intDealerDraw2 + intDealerDraw3 + intDealerDraw4 + intDealerDraw5)
        End If
        If lblDealerScore.Text < 17 Then
            Me.lblDealerCard6.Text = intDealerDraw6
            Me.lblDealerCard6.Visible = True
            Me.lblDealerScore.Text = Val(intDealerDraw1 + intDealerDraw2 + intDealerDraw3 + intDealerDraw4 + intDealerDraw5 + intDealerDraw6)
        End If
        If lblDealerScore.Text < 17 Then
            Me.lblDealerCard7.Text = intDealerDraw7
            Me.lblDealerCard7.Visible = True
            Me.lblDealerScore.Text = Val(intDealerDraw1 + intDealerDraw2 + intDealerDraw3 + intDealerDraw4 + intDealerDraw5 + intDealerDraw6 + intDealerDraw7)
        End If
        If lblDealerScore.Text < 17 Then
            Me.lblDealerCard8.Text = intDealerDraw8
            Me.lblDealerCard8.Visible = True
            Me.lblDealerScore.Text = Val(intDealerDraw1 + intDealerDraw2 + intDealerDraw3 + intDealerDraw4 + intDealerDraw5 + intDealerDraw6 + intDealerDraw7 + intDealerDraw8)
        End If
        If lblDealerScore.Text < 17 Then
            Me.lblDealerCard9.Text = intDealerDraw9
            Me.lblDealerCard9.Visible = True
            Me.lblDealerScore.Text = Val(intDealerDraw1 + intDealerDraw2 + intDealerDraw3 + intDealerDraw4 + intDealerDraw5 + intDealerDraw6 + intDealerDraw7 + intDealerDraw8 + intDealerDraw9)
        End If

        'Message
        If Me.lblPlayerScore.Text < 21 And Me.lblDealerScore.Text < 21 Then
            If Me.lblPlayerScore.Text = Me.lblDealerScore.Text Then
                Me.lblMessage.Text = "Tie! No One Wins!"
            ElseIf Me.lblDealerScore.Text > Me.lblPlayerScore.Text Then
                Me.lblMessage.Text = "Dealer Wins!"
            ElseIf lblPlayerScore.Text > lblDealerScore.Text Then
                Me.lblMessage.Text = "Player Wins!"
            End If
        End If
        If Me.lblDealerScore.Text > 21 Then
            Me.lblMessage.Text = "Dealer Busts! Player Wins!"
        End If
        If Me.lblPlayerScore.Text = 21 Then
            Me.lblMessage.Text = "BlackJack! Player Wins!"
        End If
        If Me.lblDealerScore.Text = 21 Then
            Me.lblMessage.Text = "BlackJack! Dealer Wins!"
        End If
        If Me.lblPlayerScore.Text > 21 Then
            Me.lblMessage.Text = "Player Busts! Dealer Wins!"
        End If
        If Me.lblPlayerScore.Text = 21 And Me.lblDealerScore.Text = 21 Then
            Me.lblMessage.Text = "Tie! No One Wins!"
        End If
    End Sub

    Private Sub btnHit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHit.Click
        'Clear Score
        Me.lblPlayerScore.Text = Nothing

        'Counter
        intCounter = intCounter + 1

        'Hit
        If intCounter = 1 Then
            intHitDraw3 = Int(Rnd() * 10) + 1
            lblUserCard3.Visible = True
            If intHitDraw3 = 1 Then
                Me.lblUserCard3.Text = "A"
                strAce = Val(InputBox("Ace Value 1 or 11?"))
                Me.lblUserCard3.Text = strAce
                intHitDraw3 = strAce
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3)
            Else
                Me.lblUserCard3.Text = intHitDraw3
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3)
            End If
        End If
        If intCounter = 2 Then
            intHitDraw4 = Int(Rnd() * 10) + 1
            lblUserCard4.Visible = True
            If intHitDraw4 = 1 Then
                Me.lblUserCard4.Text = "A"
                strAce = Val(InputBox("Ace Value 1 or 11?"))
                Me.lblUserCard4.Text = strAce
                intHitDraw4 = strAce
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3 + intHitDraw4)
            Else
                Me.lblUserCard4.Text = intHitDraw4
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3 + intHitDraw4)
            End If
        End If
        If intCounter = 3 Then
            intHitDraw5 = Int(Rnd() * 10) + 1
            lblUserCard5.Visible = True
            If intHitDraw5 = 1 Then
                Me.lblUserCard5.Text = "A"
                strAce = Val(InputBox("Ace Value 1 or 11?"))
                Me.lblUserCard5.Text = strAce
                intHitDraw5 = strAce
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3 + intHitDraw4 + intHitDraw5)
            Else
                Me.lblUserCard5.Text = intHitDraw5
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3 + intHitDraw4 + intHitDraw5)
            End If
        End If
        If intCounter = 4 Then
            intHitDraw6 = Int(Rnd() * 10) + 1
            lblUserCard6.Visible = True
            If intHitDraw6 = 1 Then
                Me.lblUserCard6.Text = "A"
                strAce = Val(InputBox("Ace Value 1 or 11?"))
                Me.lblUserCard6.Text = strAce
                intHitDraw6 = strAce
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3 + intHitDraw4 + intHitDraw5 + intHitDraw6)
            Else
                Me.lblUserCard6.Text = intHitDraw6
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3 + intHitDraw4 + intHitDraw5 + intHitDraw6)
            End If
        End If
        If intCounter = 5 Then
            intHitDraw7 = Int(Rnd() * 10) + 1
            lblUserCard7.Visible = True
            If intHitDraw7 = 1 Then
                Me.lblUserCard7.Text = "A"
                strAce = Val(InputBox("Ace Value 1 or 11?"))
                Me.lblUserCard7.Text = strAce
                intHitDraw7 = strAce
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3 + intHitDraw4 + intHitDraw5 + intHitDraw6 + intHitDraw7)
            Else
                Me.lblUserCard7.Text = intHitDraw7
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3 + intHitDraw4 + intHitDraw5 + intHitDraw6 + intHitDraw7)
            End If
        End If
        If intCounter = 6 Then
            intHitDraw8 = Int(Rnd() * 10) + 1
            lblUserCard8.Visible = True
            If intHitDraw8 = 1 Then
                Me.lblUserCard8.Text = "A"
                strAce = Val(InputBox("Ace Value 1 or 11?"))
                Me.lblUserCard8.Text = strAce
                intHitDraw8 = strAce
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3 + intHitDraw4 + intHitDraw5 + intHitDraw6 + intHitDraw7 + intHitDraw8)
            Else
                Me.lblUserCard8.Text = intHitDraw8
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3 + intHitDraw4 + intHitDraw5 + intHitDraw6 + intHitDraw7 + intHitDraw8)
            End If
        End If
        If intCounter = 7 Then
            intHitDraw9 = Int(Rnd() * 10) + 1
            lblUserCard9.Visible = True
            If intHitDraw9 = 1 Then
                Me.lblUserCard9.Text = "A"
                strAce = Val(InputBox("Ace Value 1 or 11?"))
                Me.lblUserCard9.Text = strAce
                intHitDraw9 = strAce
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3 + intHitDraw4 + intHitDraw5 + intHitDraw6 + intHitDraw7 + intHitDraw8 + intDealerDraw9)
            Else
                Me.lblUserCard9.Text = intHitDraw9
                Me.lblPlayerScore.Text = Val(intUserDraw1 + intUserDraw2 + intHitDraw3 + intHitDraw4 + intHitDraw5 + intHitDraw6 + intHitDraw7 + intHitDraw8 + intDealerDraw9)
            End If
        End If
        If Me.lblPlayerScore.Text > 21 Then
            Me.lblMessage.Text = "Player Busts! Dealer Wins!"
        End If
    End Sub

    Private Sub btnHelp_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnHelp.Click
        MessageBox.Show("BlackJack is a game to 21." & vbNewLine & "The player must get closer to 21 than the dealer without going over." & vbNewLine & "The Dealer must Hit until his card value is >= 17" & vbNewLine & "Face cards are worth 10" & vbNewLine & "In this version, the player choses the value of Ace's" & vbNewLine & "The Ace can either be a 1 or an 11" & vbNewLine & "The Ace card is randomized for the dealer" & vbNewLine & "Press the Shuffle button to reset the hands." & vbNewLine & "Press the Hit button to draw another card." & vbNewLine & "Press the Stand button to compare your hand to the dealer." & vbNewLine & "Have Fun!" & vbNewLine & vbNewLine & "By: Tom Szklarzewski" & vbNewLine, "How To Play BlackJack")
    End Sub
End Class
