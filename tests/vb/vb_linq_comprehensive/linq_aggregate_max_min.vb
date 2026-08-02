' vybe-test: vb/vb_linq_comprehensive/linq_aggregate_max_min
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
        Dim nums() = {15, 2, 88, 42}
        Dim mx = Aggregate n In nums Into Max()
        Dim mn = Aggregate n In nums Into Min()
        __Check(CStr(mx), "88")
        __Check(CStr(mn), "2")
    End Sub
End Module
