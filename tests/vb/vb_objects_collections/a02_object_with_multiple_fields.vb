' vybe-test: vb/vb_objects_collections/a02_object_with_multiple_fields
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

Class Person
    Public Name As String
    Public Age As Integer
    Public City As String
End Class
Dim p As New Person()
p.Name = "Alice"
p.Age = 30
p.City = "Paris"
__Check(CStr(p.Name), "Alice")
__Check(CStr(p.Age), "30")
__Check(CStr(p.City), "Paris")
