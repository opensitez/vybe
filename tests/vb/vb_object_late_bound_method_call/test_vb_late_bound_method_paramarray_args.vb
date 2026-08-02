' vybe-test: vb/vb_object_late_bound_method_call/test_vb_late_bound_method_paramarray_args
' origin: languages/vb/tests/vb/test_vb_object_late_bound_method_call.rs

Module Program
    Class MathOps
        Public Function SumAll(ParamArray numbers As Integer()) As Integer
            Dim sum = 0
            For Each n In numbers
                sum += n
            Next
            Return sum
        End Function
    End Class

    Sub Main()
        Dim obj As Object = New MathOps()
        Dim total As Integer = CInt(obj.SumAll(10, 20, 30))
        Console.WriteLine(total)
    End Sub
End Module
