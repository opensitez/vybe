' vybe-test: vb/vb_inheritance/t13_base_field_set_in_ctor_read_from_derived
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
    Public Data As String

    Sub New()
        Data = "initialized"
    End Sub
End Class

Class Child
    Inherits Base

    Function GetData() As String
        GetData = Data
    End Function
End Class

Dim c As New Child()
__Check(CStr(c.GetData()), "initialized")
