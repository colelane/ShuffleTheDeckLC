Option Strict On
Option Explicit On
Option Compare Text
'Lane Coleman
'RCET 0265
'Fall 2020
'Shuffle The Deck
'https://github.com/colelane/ShuffleTheDeckLC

Module ShuffleTheDeckLC

    Sub Main()
        Dim i As Single
        Dim j As Single
        Dim newVal As Single
        Dim cards(52) As Integer
        Dim load As Integer
        Randomize()

Shuffle:
        For load = 0 To 51
            cards(load) = load + 1
            'Console.Write(cards(load) & vbTab)
        Next

        Do
            Console.WriteLine($"Tap enter to deal cards.
Press q to quit at any time.
Press spacebar at anytime to shuffle and restart")

            Console.ReadLine()

            For i = 0 To 51
                'This section takes one of the 52 numbers in the cards array, and randomly
                'swaps it with another number.  Can't ever swap with an array cell lower than the current
                'cell, so no duplicates come up.
                j = Int((51 - i + 1) * Rnd() + i)
                newVal = cards(CInt(i))
                cards(CInt(i)) = cards(CInt(j))
                cards(CInt(j)) = CInt(newVal)

                'this entire if statement section is just for displaying the card names.
                'all of the shuffling takes place in only a few lines.
                If cards(CInt(i)) >= 2 And cards(CInt(i)) <= 10 Then
                    Console.WriteLine($"{cards(CInt(i))} of Diamonds")
                ElseIf cards(CInt(i)) = 1 Then
                    Console.WriteLine($"Ace of Diamonds")
                ElseIf cards(CInt(i)) = 11 Then
                    Console.WriteLine($"Jack of Diamonds")
                ElseIf cards(CInt(i)) = 12 Then
                    Console.WriteLine($"Queen of Diamonds")
                ElseIf cards(CInt(i)) = 13 Then
                    Console.WriteLine($"King of Diamonds")
                End If

                If cards(CInt(i)) >= 15 And cards(CInt(i)) <= 23 Then
                    Console.WriteLine($"{cards(CInt(i)) - 13} of Spades")
                ElseIf cards(CInt(i)) = 14 Then
                    Console.WriteLine($"Ace of Spades")
                ElseIf cards(CInt(i)) = 24 Then
                    Console.WriteLine($"Jack of Spades")
                ElseIf cards(CInt(i)) = 25 Then
                    Console.WriteLine($"Queen of Spades")
                ElseIf cards(CInt(i)) = 26 Then
                    Console.WriteLine($"King of Spades")
                End If

                If cards(CInt(i)) >= 28 And cards(CInt(i)) <= 36 Then
                    Console.WriteLine($"{cards(CInt(i)) - 26} of Hearts")
                ElseIf cards(CInt(i)) = 27 Then
                    Console.WriteLine($"Ace of Hearts")
                ElseIf cards(CInt(i)) = 37 Then
                    Console.WriteLine($"Jack of Hearts")
                ElseIf cards(CInt(i)) = 38 Then
                    Console.WriteLine($"Queen of Hearts")
                ElseIf cards(CInt(i)) = 39 Then
                    Console.WriteLine($"King of Hearts")
                End If

                If cards(CInt(i)) >= 41 And cards(CInt(i)) <= 49 Then
                    Console.WriteLine($"{cards(CInt(i)) - 39} of Clubs")
                ElseIf cards(CInt(i)) = 40 Then
                    Console.WriteLine($"Ace of Clubs")
                ElseIf cards(CInt(i)) = 50 Then
                    Console.WriteLine($"Jack of Clubs")
                ElseIf cards(CInt(i)) = 51 Then
                    Console.WriteLine($"Queen of Clubs")
                ElseIf cards(CInt(i)) = 52 Then
                    Console.WriteLine($"King of Clubs")
                End If


                Select Case Console.ReadKey().Key
                    Case ConsoleKey.Q
                        Exit Sub
                    Case ConsoleKey.Spacebar
                        Console.Clear()
                        GoTo Shuffle
                End Select

            Next i
        Loop Until i = 52

        Console.WriteLine("All cards have been dealt")
        Console.WriteLine("Press Space to shuffle the deck and start again. Press Q to quit")
        Do
            Select Case Console.ReadKey().Key
                Case ConsoleKey.Spacebar
                    Console.Clear()
                    GoTo Shuffle
                Case ConsoleKey.Enter
                    'This is for if the user is holding down enter.  Every once in a while,
                    'there is a strange error where the message box 'sticks' and adds
                    'many message boxes.
                    MsgBox("Woahhh! Slow down there partner!")
                Case ConsoleKey.Q
                    Exit Sub
            End Select
        Loop
    End Sub

End Module
