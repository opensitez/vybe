' vybe-test: vb/vb_convert_to_base64_string/test_vb_convert_to_base64_string_basic
' origin: languages/vb/tests/vb/test_vb_convert_to_base64_string.rs

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
Imports System.Text

Module Program
    Sub Main()
        Dim bytes As Byte() = Encoding.UTF8.GetBytes("Hello World")
        Dim base64Str = Convert.ToBase64String(bytes)
        __Check(CStr(base64Str), "SGVsbG8gV29ybGQ=")
    End Sub
End Module
