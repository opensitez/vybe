' vybe-test: vb/vb_system_bitwise_operation_matrix/bitwise_operations_integer_and_or_xor_identity
' origin: languages/vb/tests/vb/test_vb_system_bitwise_operation_matrix.rs

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
        Dim a As Integer = &HF0
        Dim b As Integer = &H0F

        __Check(CStr((a And b) = 0), "True")
        __Check(CStr((a Or b) = &HFF), "True")
        __Check(CStr((a Xor b) = &HFF), "True")
        __Check(CStr((a Xor (a Or b)) = b), "True")
    End Sub
End Module
