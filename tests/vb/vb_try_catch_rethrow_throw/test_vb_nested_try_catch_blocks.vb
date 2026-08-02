' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_nested_try_catch_blocks
' origin: languages/vb/tests/vb/test_vb_try_catch_rethrow_throw.rs

Imports System

Module Program
    Sub Main()
        Try
            Console.WriteLine("Outer Try Start")
            Try
                Throw New OverflowException("Inner Exception")
            Catch ex As OverflowException
                Console.WriteLine("Inner Catch: " & ex.Message)
            End Try
            Console.WriteLine("Outer Try End")
        Catch ex As Exception
            Console.WriteLine("Outer Catch")
        End Try
    End Sub
End Module
