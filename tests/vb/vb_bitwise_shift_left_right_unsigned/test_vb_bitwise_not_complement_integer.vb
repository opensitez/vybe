' vybe-test: vb/vb_bitwise_shift_left_right_unsigned/test_vb_bitwise_not_complement_integer
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
        Dim val As Integer = 0
        Dim comp = Not val
        __Check(CStr(comp), "-1")
    End Sub
End Module
