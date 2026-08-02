' vybe-test: vb/vb_linq_element_at_or_default/test_vb_linq_element_at_valid_and_out_of_bounds
' origin: languages/vb/tests/vb/test_vb_linq_element_at_or_default.rs

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
        Dim arr = {10, 20, 30}
        __Check(CStr(arr.ElementAt(1)), "20")
        __Check(CStr(arr.ElementAtOrDefault(5)), "0")
    End Sub
End Module
