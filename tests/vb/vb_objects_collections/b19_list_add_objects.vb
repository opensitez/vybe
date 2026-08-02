' vybe-test: vb/vb_objects_collections/b19_list_add_objects
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

Class Cat
    Public Name As String
End Class
Dim list As New List(Of Cat)
Dim c1 As New Cat()
c1.Name = "Whiskers"
Dim c2 As New Cat()
c2.Name = "Mittens"
list.Add(c1)
list.Add(c2)
__Check(CStr(list.Item(0).Name), "Whiskers")
__Check(CStr(list.Item(1).Name), "Mittens")
