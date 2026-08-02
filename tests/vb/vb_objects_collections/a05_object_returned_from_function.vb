' vybe-test: vb/vb_objects_collections/a05_object_returned_from_function
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

Class Item
    Public Label As String
End Class
Function MakeItem(lbl As String) As Item
    Dim it As New Item()
    it.Label = lbl
    Return it
End Function
Dim x As Item = MakeItem("hello")
__Check(CStr(x.Label), "hello")
