' vybe-test: vb/vb_linq_aggregate_seed_accumulator/test_vb_linq_aggregate_with_seed
' origin: languages/vb/tests/vb/test_vb_linq_aggregate_seed_accumulator.rs

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
        Dim words = {"apple", "banana", "cherry"}
        Dim totalChars = words.Aggregate(0, Function(total, nextWord) total + nextWord.Length)
        __Check(CStr(totalChars), "17")
    End Sub
End Module
