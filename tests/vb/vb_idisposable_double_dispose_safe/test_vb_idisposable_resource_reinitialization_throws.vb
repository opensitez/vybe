' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_idisposable_resource_reinitialization_throws
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

Class SingleUseResource
    Implements IDisposable
    Private isDisposed As Boolean = False

    Public Sub Initialize()
        If isDisposed Then Throw New ObjectDisposedException("SingleUseResource")
        __Check(CStr("Initialized"), "Initialized")
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        isDisposed = True
    End Sub
End Class

Module Program
    Sub Main()
        Dim res As New SingleUseResource()
        res.Initialize()
        res.Dispose()
        Try
            res.Initialize()
        Catch ex As ObjectDisposedException
            __Check(CStr("Cannot Reinitialize Disposed Resource"), "Cannot Reinitialize Disposed Resource")
        End Try
    End Sub
End Module
