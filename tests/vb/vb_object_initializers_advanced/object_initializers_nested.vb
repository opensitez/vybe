' vybe-test: vb/vb_object_initializers_advanced/object_initializers_nested
' origin: languages/vb/tests/vb/test_vb_object_initializers_advanced.rs

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
    Public Property HomeAddress As Address
End Class

Module M
    Sub Main()
        ' Nested object initializers
        Dim p As New Person() With {
            .Name = "Alice",
            .HomeAddress = New Address() With {
                .City = "New York",
                .Zip = "10001"
            }
        }
        
        __Check(CStr(p.Name), "Alice")
        __Check(CStr(p.HomeAddress.City), "New York")
    End Sub
End Module
