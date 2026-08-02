' vybe-test: vb/vb_async_configure_await_false/test_vb_async_synchronization_context_preservation
' origin: languages/vb/tests/vb/test_vb_async_configure_await_false.rs

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

Imports System.Threading
Imports System.Threading.Tasks

Class DummySyncContext
    Inherits SynchronizationContext
    Public Overrides Sub Post(d As SendOrPostCallback, state As Object)
        __Check(CStr("SyncContext Post Called"), "SyncContext Post Called")
        d(state)
    End Sub
End Class

Module Program
    Private Async Function ContextAsync() As Task
        Await Task.Yield()
    End Function

    Sub Main()
        Dim originalCtx = SynchronizationContext.Current
        Try
            SynchronizationContext.SetSynchronizationContext(New DummySyncContext())
            Dim t = ContextAsync()
            t.Wait()
        Finally
            SynchronizationContext.SetSynchronizationContext(originalCtx)
        End Try
    End Sub
End Module
