' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_handler_null_target_safe
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

Class NullTargetEmitter
    Public Event OnData As Action(Of Integer)
    Public Sub Push(v As Integer)
        RaiseEvent OnData(v)
    End Sub
End Class

Module Program
    Sub Main()
        Dim e As New NullTargetEmitter()
        Dim nullDelegate As Action(Of Integer) = Nothing
        RemoveHandler e.OnData, nullDelegate
        __Check(CStr("Safely handled null delegate removal"), "Safely handled null delegate removal")
    End Sub
End Module
