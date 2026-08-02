' vybe-test: vb/vb_interop/f46_nested_property_chain
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

Public Class Inner
    Dim value As String
    Public Sub New(v As String)
        value = v
    End Sub
End Class
Public Class Outer
    Dim inner As Inner
    Public Sub New()
        inner = New Inner("deep")
    End Sub
End Class
Dim o As New Outer()
__Check(CStr(o.inner.value), "deep")
