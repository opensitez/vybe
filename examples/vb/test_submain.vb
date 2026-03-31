Module Module1
    Sub Main()
        Console.WriteLine("Hello from Sub Main!")
        Console.WriteLine("This is a console application")
        Console.WriteLine("")
        
        ' Test the new functions
        Dim x As Double
        x = CCur(123.456789)
        Console.WriteLine("CCur(123.456789) = " & x)
        
        Dim v As Variant
        v = CVar(42)
        Console.WriteLine("CVar(42) = " & v)
        
        Console.WriteLine("")
        Console.WriteLine("App.Path = " & App.Path)
        Console.WriteLine("App.Title = " & App.Title)
        
        Console.WriteLine("")
        Console.WriteLine("Program finished")
    End Sub
End Module
