' vybe-test: vb/vb_system_array_matrix/array_sort_and_reverse_ordering
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
        Dim values() As Integer = {7, 1, 9, 2, 6}
        Array.Sort(values)
        Dim ascendingFirst As Integer = values(0)
        Dim ascendingLast As Integer = values(4)

        Array.Reverse(values)
        Dim descendingFirst As Integer = values(0)
        Dim descendingLast As Integer = values(4)

        __Check(CStr(ascendingFirst), "1")
        __Check(CStr(ascendingLast), "9")
        __Check(CStr(descendingFirst), "9")
        __Check(CStr(descendingLast), "1")
    End Sub
End Module
