Option Strict On
Option Explicit On



Module ShuffleTheDeckLC

    Sub Main()
        Dim min_item As Single
        Dim max_item As Single
        Dim i As Single
        Dim j As Single
        Dim newVal As Single
        Dim cards(52) As Integer
        Dim load As Integer

        For load = 0 To 51
            cards(load) = load + 1
            Console.Write(cards(load) & vbTab)
        Next

        Do
            min_item = LBound(cards)
            max_item = UBound(cards)

            Console.WriteLine()
            Console.ReadLine()

            For i = min_item To max_item - 1
                ' Randomly assign item number i.
                j = Int((max_item - i + 1) * Rnd() + i)
                newVal = cards(CInt(i))
                cards(CInt(i)) = cards(CInt(j))
                cards(CInt(j)) = CInt(newVal)

                Console.Write(cards(CInt(i)) & vbTab)
            Next i

        Loop
    End Sub

End Module
