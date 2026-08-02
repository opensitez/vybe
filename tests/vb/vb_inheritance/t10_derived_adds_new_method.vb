' vybe-test: vb/vb_inheritance/t10_derived_adds_new_method
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
    Function Hello() As String
        Hello = "hi"
    End Function
End Class

Class Child
    Inherits Base

    Function World() As String
        World = "world"
    End Function
End Class

Dim c As New Child()
__Check(CStr(c.Hello()), "hi")
__Check(CStr(c.World()), "world")
