' vybe-test: vb/vb_paramarray_modifier/paramarray_modifier
' origin: languages/vb/tests/vb/test_vb_paramarray_modifier.rs

Module M
    ' ParamArray allows passing a variable number of arguments
    Function Sum(ParamArray nums() As Integer) As Integer
        Dim total = 0
        For Each n In nums
            total += n
        Next
        Return total
    End Function

    Sub Main()
        Console.WriteLine(Sum())
        Console.WriteLine(Sum(1, 2, 3))
        
        Dim arr() As Integer = {4, 5, 6}
        Console.WriteLine(Sum(arr))
    End Sub
End Module
