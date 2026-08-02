' vybe-test: vb/vb_linq_skip_take_while/test_vb_linq_takewhile_none_match
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
        Dim nums = {1, 3, 5}
        Dim result = nums.TakeWhile(Function(n) n Mod 2 = 0)
        __Check(CStr(result.Count()), "0")
    End Sub
End Module
