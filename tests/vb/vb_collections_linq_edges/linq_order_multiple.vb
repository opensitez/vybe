' vybe-test: vb/vb_collections_linq_edges/linq_order_multiple
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

Imports System.Linq: Module M: Sub Main(): Dim n = {New With {.A = 1, .B = 2}, New With {.A = 1, .B = 1}}: Dim q = From x In n Order By x.A, x.B Descending Select x.B: __Check(CStr(q.First()), "2"): End Sub: End Module
