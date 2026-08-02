' vybe-test: vb/vb_procedure_arguments/arg_byref_type_coercion_copy_back
' origin: languages/vb/tests/vb/test_vb_procedure_arguments.rs

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
Sub Mutate(ByRef v As Double)
v = 2.5
End Sub
Sub Main()
Dim x As Integer = 1
Mutate(x) ' Passed as Double, copied back to Integer with conversion
__Check(CStr(x), "2")
End Sub
End Module
