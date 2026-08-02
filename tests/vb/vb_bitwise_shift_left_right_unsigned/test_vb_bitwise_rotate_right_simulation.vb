' vybe-test: vb/vb_bitwise_shift_left_right_unsigned/test_vb_bitwise_rotate_right_simulation
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
    Private Function RotateRight(val As UInteger, shift As Integer) As UInteger
        Return (val >> shift) Or (val << (32 - shift))
    End Function

    Sub Main()
        Dim val As UInteger = &H00000001UI
        Dim rot = RotateRight(val, 1)
        __Check(CStr(Hex(rot)), "80000000")
    End Sub
End Module
