' vybe-test: vb/vb_array_convertall_transformations/test_vb_array_convertall_ascii_code_to_char
' origin: languages/vb/tests/vb/test_vb_array_convertall_transformations.rs

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

Module Program
    Sub Main()
        Dim ascii As Integer() = {65, 66, 67}
        Dim chars As Char() = Array.ConvertAll(ascii, Function(i) ChrW(i))
        __Check(CStr(New String(chars)), "ABC")
    End Sub
End Module
