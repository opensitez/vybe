' vybe-test: vb/vb_custom_event_thread_sync/test_vb_custom_event_read_only_accessors
' origin: languages/vb/tests/vb/test_vb_custom_event_thread_sync.rs

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

Class ReadOnlyEventSource
    Private list As EventHandler

    Public Custom Event SimpleEvent As EventHandler
        AddHandler(value As EventHandler)
            list = CType(Delegate.Combine(list, value), EventHandler)
        End AddHandler
        RemoveHandler(value As EventHandler)
            list = CType(Delegate.Remove(list, value), EventHandler)
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            If list IsNot Nothing Then list(sender, e)
        End RaiseEvent
    End Event

    Public Sub Run()
        RaiseEvent SimpleEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim r As New ReadOnlyEventSource()
        AddHandler r.SimpleEvent, Sub(s, e) __Check(CStr("Simple Event Fired"), "Simple Event Fired")
        r.Run()
    End Sub
End Module
