' vybe-test: vb/vb_tuple_return_multiple_values/test_vb_method_returning_named_tuple
' origin: languages/vb/tests/vb/test_vb_tuple_return_multiple_values.rs

Module Program
    Function GetMinMax(numbers As Integer()) As (Min As Integer, Max As Integer)
        Dim minVal As Integer = numbers(0)
        Dim maxVal As Integer = numbers(0)
        For Each n In numbers
            If n < minVal Then minVal = n
            If n > maxVal Then maxVal = n
        Next
        Return (minVal, maxVal)
    End Function

    Sub Main()
        Dim res = GetMinMax({5, 2, 9, 1, 7})
        Console.WriteLine("Min=" & res.Min & ", Max=" & res.Max)
    End Sub
End Module
