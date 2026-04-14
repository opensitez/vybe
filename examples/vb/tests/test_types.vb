Module TypeTest
    Sub Main()
        Console.WriteLine("Type Test Start")

        ' Byte
        Dim b = CByte(255)
        Console.WriteLine("CByte(255) = " & b)
        Console.WriteLine("TypeName(b) = " & TypeName(b))

        ' Char
        Dim c = CChar("A")
        Console.WriteLine("CChar('A') = " & c)
        Console.WriteLine("TypeName(c) = " & TypeName(c))
        
        Dim c2 = CChar(65)
        Console.WriteLine("CChar(65) = " & c2)

        ' Hex
        Dim h = CInt("&HFF")
        Console.WriteLine("CInt('&HFF') = " & h)
        
        Dim h2 = CLng("&H100")
        Console.WriteLine("CLng('&H100') = " & h2)
        
        ' Octal
        Dim o = CInt("&O10") ' 8
        Console.WriteLine("CInt('&O10') = " & o)
        
        ' Overflow check (optional, but good to know)
        ' Dim err = CByte(256) ' Should error

        Console.WriteLine("Type Test Completed")
    End Sub
End Module
