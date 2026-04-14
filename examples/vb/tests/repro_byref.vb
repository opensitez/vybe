Module MainModule
    Sub Main()
        Dim x As Integer
        x = 10
        Console.WriteLine("Before: " & x)
        
        ' Should modify x to 20
        ModifyByRef(x)
        
        Console.WriteLine("After: " & x)
        
        If x = 20 Then
            Console.WriteLine("SUCCESS: ByRef works")
        Else
            Console.WriteLine("FAILURE: ByRef failed, x is " & x)
        End If
    End Sub

    Sub ModifyByRef(ByRef val As Integer)
        Console.WriteLine("Inside ModifyByRef, received: " & val)
        val = 20
        Console.WriteLine("Inside ModifyByRef, changed to: " & val)
    End Sub
End Module
