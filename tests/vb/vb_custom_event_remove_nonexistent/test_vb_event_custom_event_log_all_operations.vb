' vybe-test: vb/vb_custom_event_remove_nonexistent/test_vb_event_custom_event_log_all_operations
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

Class LoggingPublisher
    Public Custom Event CustomLog As Action(Of String)
        AddHandler(value As Action(Of String))
            __Check(CStr("Subscribed"), "Subscribed")
        End AddHandler
        RemoveHandler(value As Action(Of String))
            __Check(CStr("Unsubscribed"), "Unsubscribed")
        End RemoveHandler
        RaiseEvent(msg As String)
        End RaiseEvent
    End Event
End Class

Module Program
    Sub Main()
        Dim p As New LoggingPublisher()
        Dim h As Action(Of String) = Sub(s) End Sub
        AddHandler p.CustomLog, h
        RemoveHandler p.CustomLog, h
    End Sub
End Module
