' vybe-test: vb/vb_with_statement/with_nested
' origin: languages/vb/tests/vb/test_vb_with_statement.rs

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
    Public City As String
    Public ZipCode As Integer
End Class

Class Person
    Public Name As String
    Public Location As New Address()
End Class

Module M
    Sub Main()
        Dim p As New Person()
        With p
            .Name = "Bob"
            With .Location
                .City = "New York"
                .ZipCode = 10001
            End With
        End With
        __Check(CStr(p.Name), "Bob")
        __Check(CStr(p.Location.City), "New York")
        __Check(CStr(p.Location.ZipCode), "10001")
    End Sub
End Module
