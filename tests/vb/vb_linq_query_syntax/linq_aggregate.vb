' vybe-test: vb/vb_linq_query_syntax/linq_aggregate
' origin: languages/vb/tests/vb/test_vb_linq_query_syntax.rs

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

Module M
    Sub Main()
        Dim numbers As Integer() = {1, 2, 3, 4, 5}
        
        ' Aggregate is a distinct keyword in VB LINQ
        Dim sum = Aggregate n In numbers Into Sum()
        __Check(CStr(sum), "15")
        
        Dim max = Aggregate n In numbers Into Max()
        __Check(CStr(max), "5")
    End Sub
End Module
