' vybe-test: vb/vb_interop/b76_nested_class_in_expression
' origin: languages/vb/tests/vb/vb_interop_test.rs

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

Public Class Wrapper
    Dim val As Integer
    Public Sub New(v As Integer)
        val = v
    End Sub
    Public Function GetVal() As Integer
        Return val
    End Function
End Class
__Check(CStr(New Wrapper(42).GetVal()), "42")
