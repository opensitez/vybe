' vybe-test: vb/vb_linq_skip_take_while/test_vb_linq_skipwhile_indexed_predicate
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
        Dim numbers = {5, 10, 15, 20, 25}
        ' Skip elements while value <= index * 10 (5<=0 False, so skips nothing)
        Dim result = numbers.SkipWhile(Function(n, idx) n <= idx * 10)
        __Check(CStr(String.Join(",", result)), "5,10,15,20,25")
    End Sub
End Module
