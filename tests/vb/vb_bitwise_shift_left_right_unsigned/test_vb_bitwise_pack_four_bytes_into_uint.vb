' vybe-test: vb/vb_bitwise_shift_left_right_unsigned/test_vb_bitwise_pack_four_bytes_into_uint
' origin: languages/vb/tests/vb/test_vb_bitwise_shift_left_right_unsigned.rs

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

Module Program
    Sub Main()
        Dim b1 As Byte = &HA
        Dim b2 As Byte = &HB
        Dim b3 As Byte = &HC
        Dim b4 As Byte = &HD
        Dim packed As UInteger = (CUInt(b1) << 24) Or (CUInt(b2) << 16) Or (CUInt(b3) << 8) Or CUInt(b4)
        __Check(CStr(Hex(packed)), "A0B0C0D")
    End Sub
End Module
