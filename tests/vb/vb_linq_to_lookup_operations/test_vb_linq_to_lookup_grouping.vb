' vybe-test: vb/vb_linq_to_lookup_operations/test_vb_linq_to_lookup_grouping
' origin: languages/vb/tests/vb/test_vb_linq_to_lookup_operations.rs

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
        Dim words = {"apple", "apricot", "banana", "blueberry", "cherry"}
        Dim lookup = words.ToLookup(Function(w) w(0))

        __Check(CStr(lookup.Count), "3")
        __Check(CStr(String.Join(",", lookup("a"c))), "apple,apricot")
        __Check(CStr(String.Join(",", lookup("b"c))), "banana,blueberry")
    End Sub
End Module
