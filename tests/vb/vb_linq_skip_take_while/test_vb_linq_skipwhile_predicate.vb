' vybe-test: vb/vb_linq_skip_take_while/test_vb_linq_skipwhile_predicate
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
        Dim numbers = {1, 2, 3, 4, 5, 1, 2}
        Dim result = numbers.SkipWhile(Function(n) n < 4)
        __Check(CStr(String.Join(",", result)), "4,5,1,2")
    End Sub
End Module
