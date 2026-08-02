' vybe-test: vb/vb_system_bitwise_operation_matrix/bitwise_operations_byte_level_membership
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
        Dim mask As Byte = &HAA
        Dim hasA As Boolean = (mask And &H80) = &H80
        Dim has4 As Boolean = (mask And &H10) = &H10
        Dim onlyLow4 As Byte = mask And &HF0

        __Check(CStr(mask), "170")
        __Check(CStr(hasA), "True")
        __Check(CStr(has4), "False")
        __Check(CStr(onlyLow4), "160")
    End Sub
End Module
