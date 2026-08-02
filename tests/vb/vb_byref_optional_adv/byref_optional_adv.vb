' vybe-test: vb/vb_byref_optional_adv/byref_optional_adv
' origin: languages/vb/tests/vb/test_vb_byref_optional_adv.rs

Module M
    ' Optional ByRef is allowed, but the default value creates a temporary variable
    Sub Increment(Optional ByRef val As Integer = 10)
        val += 1
        Console.WriteLine(val)
    End Sub

    Sub Main()
        ' Passing no argument uses the default value in a temporary location
        Increment()
        
        ' Passing an argument modifies the passed variable
        Dim x As Integer = 5
        Increment(x)
        Console.WriteLine(x)
    End Sub
End Module
