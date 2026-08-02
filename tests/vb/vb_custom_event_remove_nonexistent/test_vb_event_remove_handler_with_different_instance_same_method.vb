' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_remove_handler_with_different_instance_same_method
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

Class Receiver
    Public Sub HandleEvent()
        __Check(CStr("Event Received"), "Event Received")
    End Sub
End Class

Class Emitter
    Public Event EventFired As Action
    Public Sub Fire()
        RaiseEvent EventFired()
    End Sub
End Class

Module Program
    Sub Main()
        Dim r1 As New Receiver()
        Dim r2 As New Receiver()
        Dim e As New Emitter()

        AddHandler e.EventFired, AddressOf r1.HandleEvent
        ' Attempt to remove using r2 instance
        RemoveHandler e.EventFired, AddressOf r2.HandleEvent
        e.Fire()
    End Sub
End Module
