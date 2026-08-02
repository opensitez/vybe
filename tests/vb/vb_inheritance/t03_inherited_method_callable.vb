' vybe-test: vb/vb_inheritance/t03_inherited_method_callable
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

Class Animal
    Public Name As String = "Cat"

    Function GetName() As String
        GetName = Name
    End Function
End Class

Class Cat
    Inherits Animal
End Class

Dim c As New Cat()
__Check(CStr(c.GetName()), "Cat")
