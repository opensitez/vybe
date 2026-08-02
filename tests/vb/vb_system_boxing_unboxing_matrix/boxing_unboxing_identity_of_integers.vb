' vybe-test: vb/vb_system_boxing_unboxing_matrix/boxing_unboxing_identity_of_integers
' origin: languages/vb/tests/vb/test_vb_system_boxing_unboxing_matrix.rs

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
    Sub Main()
        Dim boxed As Object = 21
        Dim value As Integer = CType(boxed, Integer)

        __Check(CStr(value), "21")
        __Check(CStr(TypeOf boxed Is Integer), "True")
    End Sub
End Module
