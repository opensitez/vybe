' vybe-test: vb/vb_async_configure_await_false/test_vb_async_configure_await_inside_loop
' origin: languages/vb/tests/vb/test_vb_async_configure_await_false.rs

Imports System.Threading.Tasks

Module Program
    Private Async Function LoopAsync() As Task(Of Integer)
        Dim total = 0
        For i As Integer = 1 To 3
            Await Task.Delay(2).ConfigureAwait(False)
            total += i
        Next
        Return total
    End Function

    Sub Main()
        Dim t = LoopAsync()
        Console.WriteLine(t.Result)
    End Sub
End Module
