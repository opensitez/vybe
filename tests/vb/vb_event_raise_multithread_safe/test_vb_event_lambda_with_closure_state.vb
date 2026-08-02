' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_lambda_with_closure_state
' origin: languages/vb/tests/vb/test_vb_event_raise_multithread_safe.rs

Imports System

Class CounterEmitter
    Public Event Counted As Action(Of Integer)
    Public Sub Tick()
        For i As Integer = 1 To 3
            RaiseEvent Counted(i)
        Next
    End Sub
End Class

Module Program
    Sub Main()
        Dim sum As Integer = 0
        Dim emitter As New CounterEmitter()
        AddHandler emitter.Counted, Sub(val) sum += val
        emitter.Tick()
        Console.WriteLine("Sum: " & sum)
    End Sub
End Module
