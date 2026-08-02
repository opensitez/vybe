' vybe-test: vb/vb_objects_collections/a03_nested_objects
' origin: languages/vb/tests/vb/vb_objects_collections_test.rs

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
End Class
Class Person
    Public Name As String
    Public Addr As Address
End Class
Dim a As New Address()
a.City = "London"
Dim p As New Person()
p.Name = "Bob"
p.Addr = a
__Check(CStr(p.Addr.City), "London")
