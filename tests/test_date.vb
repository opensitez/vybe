Module MainModule
    Sub Main()
        Console.WriteLine("Testing Date Type...")
        
        Dim d As Date
        d = Now
        Console.WriteLine("Current Date: " & d)
        
        Dim y As Integer
        y = Year(d)
        Console.WriteLine("Year: " & y)
        
        If y < 2000 Then
            Console.WriteLine("FAILURE: Year is too old")
        End If
        
        ' Test Arithmetic
        Dim tomorrow As Date
        tomorrow = d + 1.0
        Console.WriteLine("Tomorrow: " & tomorrow)
        
        If tomorrow > d Then
            Console.WriteLine("SUCCESS: Tomorrow is > Today")
        Else
            Console.WriteLine("FAILURE: Date comparison failed")
        End If
        
        ' Test diff
        Dim diff As Double
        diff = tomorrow - d
        Console.WriteLine("Difference (days): " & diff)
    End Sub
End Module
