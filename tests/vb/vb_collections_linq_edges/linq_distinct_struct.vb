' vybe-test: vb/vb_collections_linq_edges/linq_distinct_struct
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

Imports System.Linq: Structure S: Public V As Integer: End Structure: Module M: Sub Main(): Dim n = {New S With {.V = 1}, New S With {.V = 1}}: __Check(CStr(n.Distinct().Count()), "1"): End Sub: End Module
