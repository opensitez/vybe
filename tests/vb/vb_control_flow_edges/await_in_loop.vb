' vybe-test: vb/vb_control_flow_edges/await_in_loop
' origin: languages/vb/tests/vb/test_vb_control_flow_edges.rs

Imports System.Threading.Tasks

Module M
    Async Function Test() As Task
        For i = 1 To 2
            Await Task.Delay(1)
            Console.WriteLine(i)
        Next
    End Function

    Sub Main()
        Test().Wait()
    End Sub
End Module
