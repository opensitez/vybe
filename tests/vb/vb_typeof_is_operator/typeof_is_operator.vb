' vybe-test: vb/vb_typeof_is_operator/typeof_is_operator
' origin: languages/vb/tests/vb/test_vb_typeof_is_operator.rs

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
End Class

Class Dog
    Inherits Animal
End Class

Module M
    Sub Main()
        Dim d As New Dog()
        Dim a As Animal = d
        
        ' TypeOf ... Is / IsNot
        __Check(CStr(TypeOf d Is Dog), "True")
        __Check(CStr(TypeOf d Is Animal), "True")
        __Check(CStr(TypeOf d IsNot String), "True")
    End Sub
End Module
