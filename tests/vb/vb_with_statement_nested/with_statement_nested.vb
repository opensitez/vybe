' vybe-test: vb/vb_with_statement_nested/with_statement_nested
' origin: languages/vb/tests/vb/test_vb_with_statement_nested.rs

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

Class Address
    Public Property City As String
    Public Property Zip As String
End Class

Class Person
    Public Property Name As String
    Public Property Home As New Address()
End Class

Module M
    Sub Main()
        Dim p As New Person()
        
        ' Nested With statements
        With p
            .Name = "Alice"
            With .Home
                .City = "Wonderland"
                .Zip = "12345"
            End With
        End With
        
        __Check(CStr(p.Name), "Alice")
        __Check(CStr(p.Home.City), "Wonderland")
    End Sub
End Module
