' vybe-test: vb/vb_collections_linq_edges/linq_group_by_multiple_keys
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

Imports System.Linq: Module M: Sub Main(): Dim n = {New With {.A = 1, .B = 1, .V = 10}, New With {.A = 1, .B = 1, .V = 20}}: Dim q = From x In n Group By x.A, x.B Into Group Select A, B, Total = Group.Sum(Function(g) g.V): __Check(CStr(q.First().Total), "30"): End Sub: End Module
