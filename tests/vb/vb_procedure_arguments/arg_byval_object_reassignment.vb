' vybe-test: vb/vb_procedure_arguments/arg_byval_object_reassignment
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

Class C
Public V As Integer = 1
End Class
Module M
Sub Mutate(ByVal obj As C)
obj = New C()
obj.V = 2
End Sub
Sub Main()
Dim o As New C()
Mutate(o)
__Check(CStr(o.V), "1")
End Sub
End Module
