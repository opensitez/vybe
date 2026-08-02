' vybe-test: vb/vb_queue_stack/stack_pop_from_empty_throws_exception
' origin: languages/vb/tests/vb/test_vb_queue_stack.rs

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim s As New Stack(Of Integer)()
        Try
            s.Pop()
            Console.WriteLine("NoError")
        Catch ex As InvalidOperationException
            Console.WriteLine("Error")
        End Try
    End Sub
End Module
