' vybe-test: vb/vb_select_case_typeof/select_case_typeof_is
' origin: languages/vb/tests/vb/test_vb_select_case_typeof.rs

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

Class Cat
    Inherits Animal
End Class

Module M
    Sub TestType(obj As Object)
        ' VB doesn't natively support Select Case TypeOf obj Is ...
        ' We use Select Case True
        Select Case True
            Case TypeOf obj Is Dog
                __Check(CStr("Dog"), "Dog")
            Case TypeOf obj Is Cat
                __Check(CStr("Cat"), "Cat")
            Case Else
                __Check(CStr("Unknown"), "Unknown")
        End Select
    End Sub

    Sub Main()
        TestType(New Dog())
        TestType(New Cat())
        TestType(New Animal())
    End Sub
End Module
