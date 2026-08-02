' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_try_catch_in_constructor
' origin: languages/vb/tests/vb/test_vb_try_catch_rethrow_throw.rs

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

Class ConstructFailure
    Public Sub New()
        Try
            Throw New InvalidOperationException("Fail in New")
        Catch ex As Exception
            __Check(CStr("Handled in New"), "Handled in New")
        End Try
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As New ConstructFailure()
    End Sub
End Module
