' vybe-test: vb/vb_linq_aggregate/linq_aggregate
' origin: languages/vb/tests/vb/test_vb_linq_aggregate.rs

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

Imports System.Linq

Module M
    Sub Main()
        Dim numbers = {1, 2, 3, 4}
        
        Dim query = Aggregate n In numbers Into Sum(), Max(), Min()
        
        __Check(CStr(query.Sum), "10")
        __Check(CStr(query.Max), "4")
        __Check(CStr(query.Min), "1")
    End Sub
End Module
