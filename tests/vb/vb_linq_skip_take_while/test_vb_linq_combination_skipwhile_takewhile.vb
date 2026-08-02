' vybe-test: vb/vb_linq_skip_take_while/test_vb_linq_combination_skipwhile_takewhile
' origin: languages/vb/tests/vb/test_vb_linq_skip_take_while.rs

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
        Dim sequence = {1, 1, 2, 3, 5, 8, 13, 21}
        ' Skip ones, then take values under 10
        Dim subSeq = sequence.SkipWhile(Function(n) n = 1).TakeWhile(Function(n) n < 10)
        __Check(CStr(String.Join(",", subSeq)), "2,3,5,8")
    End Sub
End Module
