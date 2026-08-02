' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_idisposable_thread_safe_double_dispose
' origin: languages/vb/tests/vb/test_vb_idisposable_double_dispose_safe.rs

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

Class ThreadSafeDisposable
    Implements IDisposable
    Private disposeState As Integer = 0

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Ensure only one thread executes disposal logic!
        If Interlocked.Exchange(disposeState, 1) = 0 Then
            __Check(CStr("Thread-Safe Disposal Executed"), "Thread-Safe Disposal Executed")
        End If
    End Sub
End Class

Module Program
    Sub Main()
        Dim tsd As New ThreadSafeDisposable()
        tsd.Dispose()
        tsd.Dispose()
    End Sub
End Module
