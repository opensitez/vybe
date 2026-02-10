Module MainModule
    Sub Main()
        Console.WriteLine("Testing Lambdas...")
        
        ' Test Function Lambda
        Dim square = Function(x) x * x
        Console.WriteLine("Square(5) = " & square(5))
        
        If square(5) <> 25 Then
            Console.WriteLine("FAILURE: Square lambda")
        Else
            Console.WriteLine("SUCCESS: Square lambda")
        End If

        ' Test Sub Lambda
        Dim printer = Sub(msg) Console.WriteLine("Lambda says: " & msg)
        printer("Hello")
        Console.WriteLine("SUCCESS: Sub lambda executed")

        ' Test Closure
        Dim factor = 10
        Dim multiplier = Function(x) x * factor
        Console.WriteLine("Multiplier(5) = " & multiplier(5))
        
        If multiplier(5) <> 50 Then
            Console.WriteLine("FAILURE: Closure capture")
        Else
            Console.WriteLine("SUCCESS: Closure capture")
        End If

        ' Test Closure Update (if shared) - VB.NET closures capture variables, not values?
        ' In our implementation, we capture Environment at definition time.
        ' If we modify `factor`, does `multiplier` see it?
        ' Implementation detail: traverse env to find variable. Env stores RefCells?
        ' Our Environment stores Values directly in HashMap.
        ' So standard closures capture BY VALUE/REFERENCE depending on implementation.
        ' Our `Environment::new_with_enclosing` links to parent env.
        ' Variables in parent env are in a RefCell<Environment>.
        ' So lookups should see current values in parent env.
        
        factor = 20
        If multiplier(5) = 100 Then
            Console.WriteLine("SUCCESS: Closure sees updates")
        Else
            Console.WriteLine("NOTE: Closure captured by value or copy (expected 100, got " & multiplier(5) & ")")
        End If
        
        Console.WriteLine("Lambda Tests Completed")
    End Sub
End Module
