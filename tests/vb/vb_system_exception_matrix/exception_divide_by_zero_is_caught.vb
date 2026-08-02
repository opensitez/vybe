' vybe-test: vb/vb_system_exception_matrix/exception_divide_by_zero_is_caught
' origin: languages/vb/tests/vb/test_vb_system_exception_matrix.rs

Module M
    Sub Main()
        Try
            Dim zero As Integer = 0
            Console.WriteLine(1 \ zero)
        Catch ex As DivideByZeroException
            Console.WriteLine("zero")
        End Try
    End Sub
End Module
