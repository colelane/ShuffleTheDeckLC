Option Strict On
Option Explicit On
Option Compare Text


Module ShuffleTheDeckLC

    Sub Main()
        Dim userinput As String
        Dim min_item As Single
        Dim max_item As Single
        Dim i As Single
        Dim j As Single
        Dim newVal As Single
        Dim cards(52) As Integer
        Dim load As Integer
        Randomize()

        Console.WriteLine("Press enter to deal a card")
        Console.ReadLine()
Shuffle:
        For load = 0 To 51
            cards(load) = load + 1
            'Console.Write(cards(load) & vbTab)
        Next

        Do
            min_item = LBound(cards)
            max_item = UBound(cards)


            For i = min_item To max_item - 1
                ' Randomly assign item number i.
                j = Int((max_item - i + 1) * Rnd() + i)
                newVal = cards(CInt(i))
                cards(CInt(i)) = cards(CInt(j))
                cards(CInt(j)) = CInt(newVal)


                Console.Write(cards(CInt(i)))
                Console.ReadLine()
            Next i

        Loop Until i = 52

        Console.WriteLine("All cards have been dealt")
        Console.WriteLine("Press enter to shuffle the deck and start again. Enter Q to quit")
        userinput = Console.ReadLine()
        If userinput = "q" Then
            Exit Sub
        Else
            GoTo Shuffle
        End If
    End Sub

End Module
