' vybe-test: vb/vb_system_array_indexing_matrix/array_indexing_matrix_clear_and_copy_range
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

Imports System

Module M
    Sub Main()
        Dim values() As Integer = {1, 2, 3, 4, 5}
        Dim left() As Integer = New Integer(2) {}

        Array.Copy(values, 1, left, 0, 3)
        Array.Clear(values, 3, 2)

        Dim c1 As String = String.Join("|", left)
        Dim c2 As String = String.Join("|", values)
        __Check(CStr(c1), "2|3|4")
        __Check(CStr(c2), "1|2|3|0|0")
    End Sub
End Module
