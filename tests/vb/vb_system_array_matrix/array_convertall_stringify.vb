' vybe-test: vb/vb_system_array_matrix/array_convertall_stringify
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
        Dim values() As Integer = {1, 10}
        Dim tags() As String = Array.ConvertAll(values, Function(v As Integer) "v=" & v.ToString())

        __Check(CStr(tags.Length), "2")
        __Check(CStr(tags(0)), "v=1")
        __Check(CStr(tags(1)), "v=10")
    End Sub
End Module
