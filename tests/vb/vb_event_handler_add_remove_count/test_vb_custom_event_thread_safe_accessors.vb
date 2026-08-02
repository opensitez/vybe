' vybe-test: vb/vb_event_handler_add_remove_count/test_vb_custom_event_thread_safe_accessors
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
Imports System.Threading

Class ThreadSafeEventSource
    Private lockObj As New Object()
    Private handlers As EventHandler

    Public Custom Event SafeEvent As EventHandler
        AddHandler(value As EventHandler)
            SyncLock lockObj
                handlers = CType(Delegate.Combine(handlers, value), EventHandler)
            End SyncLock
        End AddHandler
        RemoveHandler(value As EventHandler)
            SyncLock lockObj
                handlers = CType(Delegate.Remove(handlers, value), EventHandler)
            End SyncLock
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            Dim copy As EventHandler
            SyncLock lockObj
                copy = handlers
            End SyncLock
            If copy IsNot Nothing Then copy(sender, e)
        End RaiseEvent
    End Event

    Public Sub Fire()
        RaiseEvent SafeEvent(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim src As New ThreadSafeEventSource()
        AddHandler src.SafeEvent, Sub(s, e) __Check(CStr("Thread Safe Event Fired"), "Thread Safe Event Fired")
        src.Fire()
    End Sub
End Module
