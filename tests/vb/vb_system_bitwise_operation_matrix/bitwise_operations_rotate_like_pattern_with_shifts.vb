' vybe-test: vb/vb_system_bitwise_operation_matrix/bitwise_operations_rotate_like_pattern_with_shifts
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
        Dim x As Integer = 0x12345678
        Dim highNibble As Integer = (x And &HF0000000) >> 28
        Dim lowNibble As Integer = x And &HF

        __Check(CStr(highNibble), "1")
        __Check(CStr(lowNibble), "8")
    End Sub
End Module
