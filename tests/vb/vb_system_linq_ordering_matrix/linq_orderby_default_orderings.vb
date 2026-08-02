' vybe-test: vb/vb_system_linq_ordering_matrix/linq_orderby_default_orderings
' origin: languages/vb/tests/vb/test_vb_system_linq_ordering_matrix.rs

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
        Dim values As Integer() = {5, 1, 3, 2}
        Dim ordered = values.OrderBy(Function(v) v).ToArray()
        Dim descending = values.OrderByDescending(Function(v) v).ToArray()

        __Check(CStr(String.Join(",", ordered)), "1,2,3,5")
        __Check(CStr(String.Join(",", descending)), "5,3,2,1")
    End Sub
End Module
