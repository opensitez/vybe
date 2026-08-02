' vybe-test: vb/vb_array_sort_custom_comparer/test_vb_array_sort_string_case_insensitive
' origin: languages/vb/tests/vb/test_vb_array_sort_custom_comparer.rs

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

Imports System.Collections

Module Program
    Sub Main()
        Dim words As String() = {"b", "A", "c", "B"}
        Array.Sort(words, StringComparer.OrdinalIgnoreCase)
        __Check(CStr(String.Join(",", words)), "A,b,B,c")
    End Sub
End Module
