' vybe-test: vb/vb_procedure_arguments/arg_array_dimensions_enforced
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

Module M
Sub Mutate(arr(,) As Integer)
arr(0, 0) = 2
End Sub
Sub Main()
Dim a(1, 1) As Integer
a(0, 0) = 1
Mutate(a)
__Check(CStr(a(0, 0)), "2")
End Sub
End Module
