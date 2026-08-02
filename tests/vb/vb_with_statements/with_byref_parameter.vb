' vybe-test: vb/vb_with_statements/with_byref_parameter
' origin: languages/vb/tests/vb/test_vb_with_statements.rs

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
Public V As Integer
End Class
Module M
Sub Mutate(ByRef c1 As C)
With c1
.V = 10
End With
End Sub
Sub Main()
Dim c1 As New C()
Mutate(c1)
__Check(CStr(c1.V), "10")
End Sub
End Module
