' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_using_statement_reassigning_variable_disposes_original
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

Class NamedRes
    Implements IDisposable
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        __Check(CStr("Disposed " & Name), "Inside Using Original")
    End Sub
End Class

Module Program
    Sub Main()
        Dim r As New NamedRes("Original")
        Using r
            __Check(CStr("Inside Using Original"), "Disposed Original")
        End Using
    End Sub
End Module
