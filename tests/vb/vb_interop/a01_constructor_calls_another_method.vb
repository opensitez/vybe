' vybe-test: vb/vb_interop/a01_constructor_calls_another_method
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

Public Class Foo
    Dim value As String
    Public Sub New()
        Setup()
    End Sub
    Private Sub Setup()
        value = "initialized"
    End Sub
    Public Function GetValue() As String
        Return value
    End Function
End Class
Dim f As New Foo()
__Check(CStr(f.GetValue()), "initialized")
