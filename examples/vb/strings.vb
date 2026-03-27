Module Program
    Sub Main()
        Console.WriteLine("=== VB String Functions ===")

        Dim s As String = "Hello World"

        Console.WriteLine("Original: " & s)
        Console.WriteLine("Length: " & CStr(Len(s)))
        Console.WriteLine("Upper: " & UCase(s))
        Console.WriteLine("Lower: " & LCase(s))
        Console.WriteLine("Left 5: " & Left(s, 5))
        Console.WriteLine("Right 5: " & Right(s, 5))
        Console.WriteLine("Mid 7: " & Mid(s, 7))
        Console.WriteLine("Mid 7,3: " & Mid(s, 7, 3))
        Console.WriteLine("InStr 'World': " & CStr(InStr(s, "World")))
        Console.WriteLine("Replace: " & Replace(s, "World", "VB"))
        Console.WriteLine("Trim: >" & Trim("  hello  ") & "<")
        Console.WriteLine("LTrim: >" & LTrim("  hello") & "<")
        Console.WriteLine("RTrim: >" & RTrim("hello  ") & "<")
        Console.WriteLine("Asc A: " & CStr(Asc("A")))
        Console.WriteLine("Chr 65: " & Chr(65))
    End Sub
End Module
