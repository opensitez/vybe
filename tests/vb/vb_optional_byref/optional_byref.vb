' vybe-test: vb/vb_optional_byref/optional_byref
' origin: languages/vb/tests/vb/test_vb_optional_byref.rs

Module M
    ' Optional ByRef parameter
    Sub Process(Optional ByRef val As Integer = 5)
        val += 10
        Console.WriteLine(val)
    End Sub

    Sub Main()
        Process()
        
        Dim x = 100
        Process(x)
        Console.WriteLine(x)
    End Sub
End Module
