' vybe-test: vb/vb_system_exception_matrix/exception_general_catch_always_runs_for_failure
' origin: languages/vb/tests/vb/test_vb_system_exception_matrix.rs

Module M
    Sub Main()
        Try
            Dim arr(0) As Integer
            Console.WriteLine(arr(1))
        Catch ex As Exception
            Console.WriteLine(ex.GetType().Name)
        End Try
    End Sub
End Module
