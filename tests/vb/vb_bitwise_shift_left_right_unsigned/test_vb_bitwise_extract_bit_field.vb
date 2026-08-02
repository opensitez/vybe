' vybe-test: vb/vb_bitwise_shift_left_right_unsigned/test_vb_bitwise_extract_bit_field
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
        Dim packedData As UInteger = &HAABBCCDDUI
        ' Extract second byte (CC = 204)
        Dim secondByte = (packedData >> 8) And &HFFUI
        __Check(CStr(Hex(secondByte)), "CC")
    End Sub
End Module
