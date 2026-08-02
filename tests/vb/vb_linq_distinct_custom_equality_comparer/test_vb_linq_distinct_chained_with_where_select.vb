' vybe-test: vb/vb_linq_distinct_custom_equality_comparer/test_vb_linq_distinct_chained_with_where_select
' origin: languages/vb/tests/vb/test_vb_linq_distinct_custom_equality_comparer.rs

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

Module Program
    Sub Main()
        Dim numbers = {1, 2, 3, 4, 5, 6, 7, 8}
        ' Filter evens, divide by 2, get distinct
        Dim result = numbers.Where(Function(n) n Mod 2 = 0).Select(Function(n) n \ 2).Distinct()
        __Check(CStr(String.Join(",", result)), "1,2,3,4")
    End Sub
End Module
