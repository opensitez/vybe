' vybe-test: vb/vb_convert_to_base64_string/test_vb_convert_from_base64_null_throws_argument_null
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

Module Program
    Sub Main()
        Try
            Convert.FromBase64String(Nothing)
        Catch ex As ArgumentNullException
            __Check(CStr("ArgumentNullException Caught on Null Base64"), "ArgumentNullException Caught on Null Base64")
        End Try
    End Sub
End Module
