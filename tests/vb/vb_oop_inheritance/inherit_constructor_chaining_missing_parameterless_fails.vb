' vybe-test: vb/vb_oop_inheritance/inherit_constructor_chaining_missing_parameterless_fails
' origin: languages/vb/tests/vb/test_vb_oop_inheritance.rs

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

Class B
Public Sub New(v As Integer)
End Sub
End Class
Class C
Inherits B
' Public Sub New() ' Fails if it doesn't explicitly call MyBase.New(v) because B has no parameterless Sub New
End Class
Module M
Sub Main()
__Check(CStr("Parsed"), "Parsed")
End Sub
End Module
