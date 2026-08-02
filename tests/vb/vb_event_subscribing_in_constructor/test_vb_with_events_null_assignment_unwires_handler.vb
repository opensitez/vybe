' vybe-test: vb/vb_event_subscribing_in_constructor/test_vb_with_events_null_assignment_unwires_handler
' origin: languages/vb/tests/vb/test_vb_event_subscribing_in_constructor.rs

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

Class Emitter
    Public Event Action As EventHandler
    Public Sub Fire()
        RaiseEvent Action(Me, EventArgs.Empty)
    End Sub
End Class

Class NullableListener
    Private WithEvents myEmitter As Emitter

    Public Sub Bind(e As Emitter)
        myEmitter = e
    End Sub

    Private Sub OnAction(sender As Object, e As EventArgs) Handles myEmitter.Action
        __Check(CStr("Fired"), "Fired")
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Emitter()
        Dim listener As New NullableListener()
        listener.Bind(e)
        e.Fire()

        listener.Bind(Nothing) ' Unwires myEmitter!
        e.Fire()
    End Sub
End Module
