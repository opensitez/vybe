Module TestNewFunctions
    Sub Main()
        ' Test CCur conversion
        Dim amount As Double
        amount = CCur(123.456789)
        Console.WriteLine("CCur(123.456789) = " & amount)
        
        ' Test CVar conversion
        Dim v As Variant
        v = CVar(42)
        Console.WriteLine("CVar(42) = " & v)
        
        ' Test IsNull
        Dim obj As Object
        obj = Nothing
        Console.WriteLine("IsNull(Nothing) = " & IsNull(obj))
        Console.WriteLine("IsNull(42) = " & IsNull(42))
        
        ' Test Erase on array
        Dim arr(5) As Integer
        arr(0) = 10
        arr(1) = 20
        Console.WriteLine("Array before Erase: " & arr(0) & ", " & arr(1))
        arr = Erase(arr)
        Console.WriteLine("Array after Erase length: " & arr.Length)
        
        ' Test App object
        Console.WriteLine("App.Path = " & App.Path)
        Console.WriteLine("App.Title = " & App.Title)
        Console.WriteLine("App.EXEName = " & App.EXEName)
        
        ' Test Screen object
        Console.WriteLine("Screen.Width = " & Screen.Width)
        Console.WriteLine("Screen.Height = " & Screen.Height)
        
        Console.WriteLine("")
        Console.WriteLine("All new functions tested successfully!")
    End Sub
End Module
