' vybe-test: vb/vb_collections_linq_edges/vb_collection_remove
' origin: languages/vb/tests/vb/test_vb_collections_linq_edges.rs

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

Module M: Sub Main(): Dim c As New Microsoft.VisualBasic.Collection(): c.Add("A", "K"): c.Remove("K"): __Check(CStr(c.Count), "0"): End Sub: End Module
