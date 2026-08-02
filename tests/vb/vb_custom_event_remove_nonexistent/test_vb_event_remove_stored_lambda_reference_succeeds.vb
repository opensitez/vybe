' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_stored_lambda_reference_succeeds
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

Class Emitter
    Public Event Message As Action(Of String)
    Public Sub Dispatch(m As String)
        RaiseEvent Message(m)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New Emitter()
        Dim handler As Action(Of String) = Sub(m) __Check(CStr("Msg: " & m), "Msg: First")
        AddHandler e.Message, handler
        e.Dispatch("First")
        RemoveHandler e.Message, handler
        e.Dispatch("Second")
    End Sub
End Module
