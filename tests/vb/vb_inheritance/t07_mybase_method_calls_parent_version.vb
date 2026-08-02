' vybe-test: vb/vb_inheritance/t07_mybase_method_calls_parent_version
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
    Sub Greet()
        __Check(CStr("hello from base"), "hello from base")
    End Sub
End Class

Class Child
    Inherits Base

    Sub Greet()
        MyBase.Greet()
        __Check(CStr("hello from child"), "hello from child")
    End Sub
End Class

Dim c As New Child()
c.Greet()
