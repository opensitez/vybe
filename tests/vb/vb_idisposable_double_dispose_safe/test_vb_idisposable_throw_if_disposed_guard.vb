' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_idisposable_throw_if_disposed_guard
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

Class GuardedResource
    Implements IDisposable
    Private isDisposed As Boolean = False

    Public Sub DoWork()
        If isDisposed Then Throw New ObjectDisposedException(NameOf(GuardedResource))
        __Check(CStr("Work Done"), "Work Done")
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        isDisposed = True
    End Sub
End Class

Module Program
    Sub Main()
        Dim res As New GuardedResource()
        res.DoWork()
        res.Dispose()

        Try
            res.DoWork()
        Catch ex As ObjectDisposedException
            __Check(CStr("ObjectDisposedException Caught"), "ObjectDisposedException Caught")
        End Try
    End Sub
End Module
