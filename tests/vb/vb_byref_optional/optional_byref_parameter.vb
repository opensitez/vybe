' vybe-test: vb/vb_byref_optional/optional_byref_parameter
' origin: languages/vb/tests/vb/test_vb_byref_optional.rs

Module M
    ' ByRef parameter can be Optional with a default value
    Sub UpdateValue(Optional ByRef val As Integer = 5)
        val += 10
        Console.WriteLine("Inside: " & val.ToString())
    End Sub

    Sub Main()
        Dim x As Integer = 2
        ' Passing explicitly (mutates x)
        UpdateValue(x)
        Console.WriteLine("After: " & x.ToString())
        
        ' Omitting creates a temporary variable initialized to 5
        UpdateValue()
    End Sub
End Module
