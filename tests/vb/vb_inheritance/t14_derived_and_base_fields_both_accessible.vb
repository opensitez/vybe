' vybe-test: vb/vb_inheritance/t14_derived_and_base_fields_both_accessible
' origin: languages/vb/tests/vb/vb_inheritance_test.rs

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

Class Base
    Public X As Integer = 10
End Class

Class Child
    Inherits Base
    Public Y As Integer = 20
End Class

Dim c As New Child()
__Check(CStr(c.X), "10")
__Check(CStr(c.Y), "20")
