' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_try_catch_exit_try_statement
' origin: languages/vb/tests/vb/test_vb_try_catch_rethrow_throw.rs

Module Program
    Sub Main()
        Try
            Console.WriteLine("Before Exit Try")
            Exit Try
            Console.WriteLine("After Exit Try")
        Catch ex As Exception
            Console.WriteLine("Catch Block")
        Finally
            Console.WriteLine("Finally Block")
        End Try
    End Sub
End Module
