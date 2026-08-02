' vybe-test: vb/vb_linq_comprehensive/linq_aggregate_sum
' origin: languages/vb/tests/vb/test_vb_linq_comprehensive.rs

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
        Dim nums() = {10, 20, 30}
        Dim sum = Aggregate n In nums Into Sum()
        __Check(CStr(sum), "60")
    End Sub
End Module
