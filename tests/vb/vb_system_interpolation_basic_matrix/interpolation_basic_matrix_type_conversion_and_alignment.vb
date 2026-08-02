' vybe-test: vb/vb_system_interpolation_basic_matrix/interpolation_basic_matrix_type_conversion_and_alignment
' origin: languages/vb/tests/vb/test_vb_system_interpolation_basic_matrix.rs

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
        Dim x As Double = 1.2
        Dim s1 As String = $"{x,8:F1}"
        Dim s2 As String = $"{x,8:F1}!"
        Dim n As Integer = 42
        Dim s3 As String = $"[{n,5}]"

        __Check(CStr(s1.Length), "8")
        __Check(CStr(s2), "   1.2!")
        __Check(CStr(s3), "[   42]")
    End Sub
End Module
