' vybe-test: vb/vb_system_array_indexing_matrix/array_indexing_matrix_value_assignment_aliases_reference
' origin: languages/vb/tests/vb/test_vb_system_array_indexing_matrix.rs

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
        Dim source() As Integer = {1, 2, 3}
        Dim alias() As Integer = source

        alias(1) = 42
        __Check(CStr(source(1)), "42")
        __Check(CStr(alias.Length), "3")
    End Sub
End Module
