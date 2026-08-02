' vybe-test: vb/vb_event_raise_multithread_safe/test_vb_event_custom_event_thread_sync_lock
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

Class ThreadSafeEventPublisher
    Private syncObj As New Object()
    Private internalDelegate As Action

    Public Custom Event SecureEvent As Action
        AddHandler(value As Action)
            SyncLock syncObj
                internalDelegate = CType([Delegate].Combine(internalDelegate, value), Action)
            End SyncLock
        End AddHandler
        RemoveHandler(value As Action)
            SyncLock syncObj
                internalDelegate = CType([Delegate].Remove(internalDelegate, value), Action)
            End SyncLock
        End RemoveHandler
        RaiseEvent()
            Dim copy As Action = Nothing
            SyncLock syncObj
                copy = internalDelegate
            End SyncLock
            If copy IsNot Nothing Then copy()
        End RaiseEvent
    End Event

    Public Sub Run()
        RaiseEvent SecureEvent()
    End Sub
End Class

Module Program
    Sub Main()
        Dim pub As New ThreadSafeEventPublisher()
        AddHandler pub.SecureEvent, Sub() __Check(CStr("ThreadSafe Event Fired"), "ThreadSafe Event Fired")
        pub.Run()
    End Sub
End Module
