' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_event_handler_lambda_closure_state_mutation
' origin: languages/vb/tests/vb/test_vb_event_handler_add_remove_count.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System

Class CounterEmitter
    Public Event Increment As EventHandler
    Public Sub Fire()
        RaiseEvent Increment(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim count = 0
        Dim emitter As New CounterEmitter()
        AddHandler emitter.Increment, Sub(s, e) count += 10
        emitter.Fire()
        emitter.Fire()
        __Check(CStr("Final Count: " & count), "Final Count: 20")
    End Sub
End Module
