' vybe-test: vb/vb_system_string_api_matrix/string_api_matrix_copy_and_padding
' origin: languages/vb/tests/vb/test_vb_system_string_api_matrix.rs

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
        Dim left As String = "abc"
        Dim right As String = "def"
        Dim padded As String = left.PadLeft(5, "_")
        Dim chars() As Char = {"a"c, "b"c, "c"c}
        Dim copied As String = New String(chars)

        __Check(CStr(left = copied), "True")
        __Check(CStr(left.PadRight(6, "-")), "abc---")
        __Check(CStr(padded), "__abc")
    End Sub
End Module
