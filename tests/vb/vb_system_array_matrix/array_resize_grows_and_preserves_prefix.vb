' vybe-test: vb/vb_system_array_matrix/array_resize_grows_and_preserves_prefix
' origin: languages/vb/tests/vb/test_vb_system_array_matrix.rs

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
        Dim values() As Integer = {1, 2, 3}
        Array.Resize(values, 5)

        values(3) = 4
        values(4) = 5

        __Check(CStr(values.Length), "5")
        __Check(CStr(values(0)), "1")
        __Check(CStr(values(4)), "5")
    End Sub
End Module
