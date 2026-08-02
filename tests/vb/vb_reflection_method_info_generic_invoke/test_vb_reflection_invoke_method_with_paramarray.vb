' vybe-test: vb/vb_reflection_method_info_generic_invoke/test_vb_reflection_invoke_method_with_paramarray
' origin: languages/vb/tests/vb/test_vb_reflection_method_info_generic_invoke.rs

Class Aggregator
    Public Function SumAll(ParamArray numbers As Integer()) As Integer
        Dim sum = 0
        For Each n In numbers : sum += n : Next
        Return sum
    End Function
End Class

Module Program
    Sub Main()
        Dim agg As New Aggregator()
        Dim m = GetType(Aggregator).GetMethod("SumAll")
        Dim res = m.Invoke(agg, {New Integer() {10, 20, 30}})
        Console.WriteLine(res)
    End Sub
End Module
