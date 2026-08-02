' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_handler_from_static_event
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

Class SharedPublisher
    Public Shared Event SharedEvent As Action
    Public Shared Sub Fire()
        RaiseEvent SharedEvent()
    End Sub
End Class

Module Program
    Private Sub OnShared() : __Check(CStr("Shared Fired"), "Shared Fired") : End Sub

    Sub Main()
        AddHandler SharedPublisher.SharedEvent, AddressOf OnShared
        SharedPublisher.Fire()
        RemoveHandler SharedPublisher.SharedEvent, AddressOf OnShared
        SharedPublisher.Fire()
    End Sub
End Module
