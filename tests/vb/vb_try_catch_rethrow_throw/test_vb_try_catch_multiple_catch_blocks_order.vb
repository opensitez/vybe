' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_try_catch_multiple_catch_blocks_order
' origin: languages/vb/tests/vb/test_vb_try_catch_rethrow_throw.rs

Imports System

Module Program
    Sub Main()
        Try
            Dim arr As Integer() = {1, 2}
            Console.WriteLine(arr(5))
        Catch ex As IndexOutOfRangeException
            Console.WriteLine("Caught IndexOutOfRangeException")
        Catch ex As Exception
            Console.WriteLine("Caught General Exception")
        End Try
    End Sub
End Module
