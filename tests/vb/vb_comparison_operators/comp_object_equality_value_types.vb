' vybe-test: vb/vb_comparison_operators/comp_object_equality_value_types
' origin: languages/vb/tests/vb/test_vb_comparison_operators.rs

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

Option Strict Off
Module M
Sub Main()
Dim obj1 As Object = 10
Dim obj2 As Object = 10
__Check(CStr(obj1 = obj2), "True")
End Sub
End Module
