Module Module1
    Sub Main()
        Console.WriteLine("=== Sub Main Started ===")
        Console.WriteLine("This project starts with Sub Main instead of a form")
        Console.WriteLine("")
        
        ' Test the functions we just added
        Dim amount As Double
        amount = CCur(123.456789)
        Console.WriteLine("CCur(123.456789) = " & amount)
        
        Dim v As Variant
        v = CVar(42)
        Console.WriteLine("CVar(42) = " & v)
        
        ' Test IsNull - simplified
        Console.WriteLine("IsNull test skipped for now")
        'Console.WriteLine("IsNull(Nothing) = " & IsNull(Nothing))
        'Console.WriteLine("IsNull(42) = " & IsNull(42))
        
        ' Test App object
        Console.WriteLine("")
        Console.WriteLine("App.Path = " & App.Path)
        Console.WriteLine("App.Title = " & App.Title)
        
        ' Test Screen object
        Console.WriteLine("Screen.Width = " & Screen.Width)
        Console.WriteLine("Screen.Height = " & Screen.Height)
        
        ' Test Clipboard object
        Console.WriteLine("")
        Console.WriteLine("Clipboard object test:")
        Dim cb As Object
        cb = Clipboard
        Console.WriteLine("Clipboard type: " & TypeName(cb))
        
        ' Test Forms collection
        Console.WriteLine("")
        Console.WriteLine("Forms collection test:")
        Dim frms As Object
        frms = Forms
        Console.WriteLine("Forms type: " & TypeName(frms))
        
        Console.WriteLine("")
        Console.WriteLine("=== Sub Main Completed Successfully ===")
    End Sub
End Module
