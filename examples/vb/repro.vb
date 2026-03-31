Module Repro
    Sub Main()
        Dim s = "hello vybe"
        Console.WriteLine("Length: " & s.Length)
        For i = 0 To s.Length - 1
            Console.WriteLine("Char " & i & ": " & Asc(s.Chars(i)))
        Next
    End Sub
End Module
