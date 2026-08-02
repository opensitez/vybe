' vybe-test: vb/vb_system_array_2d_matrix/array_2d_matrix_for_each_like_iteration_not_supported
' origin: languages/vb/tests/vb/test_vb_system_array_2d_matrix.rs

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
        Dim matrix As Integer(,) = New Integer(,) {{1, 2, 3}, {4, 5, 6}}
        Dim first As Integer = matrix(0, 0)
        Dim last As Integer = matrix(1, 2)
        __Check(CStr(first), "1")
        __Check(CStr(last), "6")
    End Sub
End Module
