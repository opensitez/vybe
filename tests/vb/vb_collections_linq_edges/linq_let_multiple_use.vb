' vybe-test: vb/vb_collections_linq_edges/linq_let_multiple_use
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

Imports System.Linq: Module M: Sub Main(): Dim n = {1}: Dim q = From x In n Let y = x + 1, z = y + 1 Select x + y + z: __Check(CStr(q.First()), "6"): End Sub: End Module
