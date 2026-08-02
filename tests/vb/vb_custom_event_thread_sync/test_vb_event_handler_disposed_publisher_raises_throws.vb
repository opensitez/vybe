' vybe-test: vb/vb_custom_event_thread_sync/test_vb_event_handler_disposed_publisher_raises_throws
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

Class DisposablePublisher
    Implements IDisposable
    Public Event DataEvent As EventHandler
    Private isDisposed As Boolean = False

    Public Sub Fire()
        If isDisposed Then Throw New ObjectDisposedException("DisposablePublisher")
        RaiseEvent DataEvent(Me, EventArgs.Empty)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        isDisposed = True
    End Sub
End Class

Module Program
    Sub Main()
        Dim dp As New DisposablePublisher()
        dp.Dispose()
        Try
            dp.Fire()
        Catch ex As ObjectDisposedException
            __Check(CStr("ObjectDisposedException Caught on Fire"), "ObjectDisposedException Caught on Fire")
        End Try
    End Sub
End Module
