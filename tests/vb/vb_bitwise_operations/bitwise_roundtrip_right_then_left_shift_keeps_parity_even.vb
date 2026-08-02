' vybe-test: vb/vb_bitwise_operations/bitwise_roundtrip_right_then_left_shift_keeps_parity_even
' origin: languages/vb/tests/vb/test_vb_bitwise_operations.rs

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
        Dim value As Integer = 9
        value = value << 2
        value = value >> 2
        __Check(CStr(value = 9), "True")
    End Sub
End Module
