' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_idisposable_check_disposed_property
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

Class ReadableStateDisposable
    Implements IDisposable

    Public ReadOnly Property IsDisposed As Boolean

    Public Sub Dispose() Implements IDisposable.Dispose
        _IsDisposed = True
        __Check(CStr("Disposed State Flag Set"), "False")
    End Sub
End Class

Module Program
    Sub Main()
        Dim r As New ReadableStateDisposable()
        __Check(CStr(r.IsDisposed), "Disposed State Flag Set")
        r.Dispose()
        __Check(CStr(r.IsDisposed), "True")
    End Sub
End Module
