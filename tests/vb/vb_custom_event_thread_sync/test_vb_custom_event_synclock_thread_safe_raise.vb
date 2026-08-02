' vybe-test: vb/vb_custom_event_thread_sync/test_vb_custom_event_synclock_thread_safe_raise
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
Imports System.Threading

Class ThreadSafeNotifier
    Private lockObj As New Object()
    Private handlerDelegate As EventHandler

    Public Custom Event StatusUpdate As EventHandler
        AddHandler(value As EventHandler)
            SyncLock lockObj
                handlerDelegate = CType(Delegate.Combine(handlerDelegate, value), EventHandler)
            End SyncLock
        End AddHandler
        RemoveHandler(value As EventHandler)
            SyncLock lockObj
                handlerDelegate = CType(Delegate.Remove(handlerDelegate, value), EventHandler)
            End SyncLock
        End RemoveHandler
        RaiseEvent(sender As Object, e As EventArgs)
            Dim temp As EventHandler
            SyncLock lockObj
                temp = handlerDelegate
            End SyncLock
            If temp IsNot Nothing Then temp(sender, e)
        End RaiseEvent
    End Event

    Public Sub Signal()
        RaiseEvent StatusUpdate(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim notifier As New ThreadSafeNotifier()
        AddHandler notifier.StatusUpdate, Sub(s, e) __Check(CStr("ThreadSafe Update Received"), "ThreadSafe Update Received")
        notifier.Signal()
    End Sub
End Module
