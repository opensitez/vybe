' vybe-test: vb/vb_linq_distinct_custom_equality_comparer/test_vb_linq_distinct_string_comparer_ignore_case
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

Imports System
Imports System.Linq

Module Program
    Sub Main()
        Dim words = {"apple", "APPLE", "banana", "Apple"}
        Dim unique = words.Distinct(StringComparer.OrdinalIgnoreCase)
        __Check(CStr(String.Join(",", unique)), "apple,banana")
    End Sub
End Module
