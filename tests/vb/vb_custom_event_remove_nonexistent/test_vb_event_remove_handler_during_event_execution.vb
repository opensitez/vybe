' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_handler_during_event_execution
' origin: languages/vb/tests/vb/test_vb_custom_event_remove_nonexistent.rs

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

Class IterativeEmitter
    Public Event StepEvent As Action
    Public Sub Fire()
        RaiseEvent StepEvent()
    End Sub
End Class

Module Program
    Private h1 As Action
    Private e As New IterativeEmitter()

    Sub Main()
        h1 = Sub()
            __Check(CStr("H1 Executing & Unsubscribing"), "H1 Executing & Unsubscribing")
            RemoveHandler e.StepEvent, h1
        End Sub

        AddHandler e.StepEvent, h1
        e.Fire()
        __Check(CStr("Second Fire:"), "Second Fire:")
        e.Fire()
    End Sub
End Module
