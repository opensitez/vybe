Module StringTest
    Sub Main()
        Console.WriteLine("String Test Start")

        ' Replace
        Dim s = "Hello World"
        Dim r = Replace(s, "World", "Universe")
        Console.WriteLine("Replace: " & r)

        ' Split & Join
        Dim csv = "A,B,C"
        Dim parts = Split(csv, ",")
        Console.WriteLine("Split(0): " & parts(0))
        Console.WriteLine("Split(1): " & parts(1))
        
        Dim joined = Join(parts, "-")
        Console.WriteLine("Join: " & joined)

        ' StrReverse
        Dim rev = StrReverse("ABC")
        Console.WriteLine("StrReverse: " & rev)

        ' InStrRev
        ' "Hello World"
        ' 12345678901
        ' l is at 3, 4, 10
        ' InStrRev("Hello World", "l") -> 10
        Dim idx = InStrRev("Hello World", "l")
        Console.WriteLine("InStrRev: " & idx)

        ' Space & String
        Dim sp = Space(3)
        Console.WriteLine("Space: '" & sp & "'")
        
        Dim st = String(3, "*")
        Console.WriteLine("String: " & st)

        ' Asc & Chr
        Dim code = Asc("A")
        Console.WriteLine("Asc: " & code)
        
        Dim ch = Chr(66)
        Console.WriteLine("Chr: " & ch)

        Console.WriteLine("String Test Completed")
    End Sub
End Module
