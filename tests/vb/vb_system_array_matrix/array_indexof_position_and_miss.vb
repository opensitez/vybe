' vybe-test: vb/vb_system_array_matrix/array_indexof_position_and_miss
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
        Dim values() As Integer = {4, 5, 6, 5}
        Dim firstFive As Integer = Array.IndexOf(values, 5)
        Dim secondFive As Integer = Array.IndexOf(values, 5, firstFive + 1)
        Dim missing As Integer = Array.IndexOf(values, 99)

        __Check(CStr(firstFive), "1")
        __Check(CStr(secondFive), "3")
        __Check(CStr(missing), "-1")
    End Sub
End Module
