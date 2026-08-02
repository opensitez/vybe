' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_idisposable_exception_in_dispose_handled
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

Class ThrowingDisposable
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Throw New InvalidOperationException("Dispose Failed")
    End Sub
End Class

Module Program
    Sub Main()
        Try
            Using res As New ThrowingDisposable()
                __Check(CStr("Inside Using Block"), "Inside Using Block")
            End Using
        Catch ex As InvalidOperationException
            __Check(CStr(ex.Message), "Dispose Failed")
        End Try
    End Sub
End Module
