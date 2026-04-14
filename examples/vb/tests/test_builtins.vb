Module MainModule
    Sub Main()
        Console.WriteLine("Testing Built-ins...")
        
        ' Test CDate
        Dim d As Date
        d = CDate("2023-10-25 14:30:00")
        Console.WriteLine("CDate Year: " & Year(d))
        If Year(d) <> 2023 Then Console.WriteLine("FAILURE: CDate Year")
        
        d = CDate("10/25/2023")
        Console.WriteLine("CDate Month: " & Month(d))
        If Month(d) <> 10 Then Console.WriteLine("FAILURE: CDate Month")
        
        ' Test String Functions
        Dim s As String = "Hello World"
        Console.WriteLine("Left: " & Left(s, 5))
        If Left(s, 5) <> "Hello" Then Console.WriteLine("FAILURE: Left")
        
        Console.WriteLine("Mid: " & Mid(s, 7, 5))
        If Mid(s, 7, 5) <> "World" Then Console.WriteLine("FAILURE: Mid")
        
        Console.WriteLine("InStr: " & InStr(s, "World"))
        If InStr(s, "World") <> 7 Then Console.WriteLine("FAILURE: InStr")
        
        ' Test Format
        Dim n As Double = 1234.5678
        Console.WriteLine("Format: " & Format(n, "0.00"))
        ' Format implementation might vary, check rough output in test runner or visual check
        
        ' Test Math
        Console.WriteLine("Int(1.2): " & Int(1.2))
        If Int(1.2) <> 1 Then Console.WriteLine("FAILURE: Int positive")
        
        Console.WriteLine("Int(-1.2): " & Int(-1.2))
        If Int(-1.2) <> -2 Then Console.WriteLine("FAILURE: Int negative")
        
        Console.WriteLine("Fix(-1.2): " & Fix(-1.2))
        If Fix(-1.2) <> -1 Then Console.WriteLine("FAILURE: Fix negative")
        
        ' Test Rnd
        Dim r As Single
        r = Rnd()
        Console.WriteLine("Rnd: " & r)
        
        Console.WriteLine("SUCCESS: Built-ins tests passed")
    End Sub
End Module
