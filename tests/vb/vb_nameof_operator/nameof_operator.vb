' vybe-test: vb/vb_nameof_operator/nameof_operator
' origin: languages/vb/tests/vb/test_vb_nameof_operator.rs

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
    Public Property Name As String
End Class

Module M
    Sub Main()
        ' NameOf returns the name of a variable, type, or member
        __Check(CStr(NameOf(Person)), "Person")
        __Check(CStr(NameOf(Person.Name)), "Name")
        
        Dim i = 10
        __Check(CStr(NameOf(i)), "i")
    End Sub
End Module
