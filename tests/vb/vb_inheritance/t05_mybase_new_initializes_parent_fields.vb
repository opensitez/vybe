' vybe-test: vb/vb_inheritance/t05_mybase_new_initializes_parent_fields
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
    Public X As Integer = 42
End Class

Class Child
    Inherits Base

    Sub New()
        MyBase.New()
    End Sub
End Class

Dim c As New Child()
__Check(CStr(c.X), "42")
