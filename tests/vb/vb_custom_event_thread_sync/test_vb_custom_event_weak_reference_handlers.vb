' vybe-test: vb/vb_custom_event_thread_sync/test_vb_custom_event_weak_reference_handlers
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

Class WeakEventSubscriber
    Public Sub OnNotify(sender As Object, e As EventArgs)
        __Check(CStr("Weak Handler Triggered"), "Weak Handler Triggered")
    End Sub
End Class

Class WeakPublisher
    Public Event Notify As EventHandler
    Public Sub Fire()
        RaiseEvent Notify(Me, EventArgs.Empty)
    End Sub
End Class

Module Program
    Sub Main()
        Dim pub As New WeakPublisher()
        Dim subObj As New WeakEventSubscriber()
        AddHandler pub.Notify, AddressOf subObj.OnNotify
        pub.Fire()
    End Sub
End Module
