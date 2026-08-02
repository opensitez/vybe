' vybe-test: vb/vb_system_integer_arithmetic_matrix/integer_arithmetic_matrix_type_boundary_checks
' origin: languages/vb/tests/vb/test_vb_system_integer_arithmetic_matrix.rs

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
        Dim maxInt As Integer = Integer.MaxValue
        Dim minInt As Integer = Integer.MinValue

        __Check(CStr(maxInt > minInt), "True")
        __Check(CStr(Integer.MinValue < 0), "True")
        __Check(CStr(CLng(maxInt) + 1 > maxInt), "True")
        __Check(CStr(CLng(maxInt) + 1 - 1 = CLng(maxInt)), "True")
    End Sub
End Module
