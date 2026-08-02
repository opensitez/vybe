' vybe-test: vb/vb_linq_all_any_predicates/test_vb_linq_any_dictionary_kvp_matching
' origin: languages/vb/tests/vb/test_vb_linq_all_any_predicates.rs

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

Imports System.Collections.Generic
Imports System.Linq

Module Program
    Sub Main()
        Dim dict As New Dictionary(Of String, Integer) From {{"A", 10}, {"B", 20}}
        __Check(CStr(dict.Any(Function(kv) kv.Value > 15)), "True")
    End Sub
End Module
