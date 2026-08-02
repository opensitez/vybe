' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_handler_exception_handling
' origin: languages/vb/tests/vb/test_vb_event_raise_multithread_safe.rs

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

Class FaultyEmitter
    Public Event Notify As Action
    Public Sub Fire()
        Try
            RaiseEvent Notify()
        Catch ex As Exception
            __Check(CStr("Caught Exception: " & ex.Message), "Caught Exception: Handler Failed")
        End Try
    End Sub
End Class

Module Program
    Sub Main()
        Dim fe As New FaultyEmitter()
        AddHandler fe.Notify, Sub() Throw New InvalidOperationException("Handler Failed")
        fe.Fire()
    End Sub
End Module
