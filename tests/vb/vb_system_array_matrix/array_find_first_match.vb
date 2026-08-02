' vybe-test: vb/vb_system_array_matrix/array_find_first_match
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
        Dim values() As Integer = {3, 5, 7, 9}
        Dim firstBig As Integer = Array.Find(values, Function(v As Integer) v > 6)
        Dim firstTiny As Integer = Array.Find(values, Function(v As Integer) v < 0)

        __Check(CStr(firstBig), "7")
        __Check(CStr(firstTiny = 0), "True")
    End Sub
End Module
