' vybe-test: vb/vb_objects_collections/b20_list_count_after_add_remove
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

Dim list As New List(Of String)
list.Add("a")
list.Add("b")
list.Add("c")
list.Add("d")
list.Remove("b")
list.Remove("d")
__Check(CStr(list.Count), "2")
