' vybe-test: vb/vb_inheritance/t19_property_get_set_on_base
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

Class Person
    Private _name As String = ""

    Property Name() As String
        Get
            Name = _name
        End Get
        Set(value As String)
            _name = value
        End Set
    End Property
End Class

Dim p As New Person()
p.Name = "Alice"
__Check(CStr(p.Name), "Alice")
