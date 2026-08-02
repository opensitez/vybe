' vybe-test: vb/vb_convert_to_base64_string/test_vb_convert_to_base64_char_array
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
        Dim input As Byte() = {1, 2, 3}
        Dim outChars(10) As Char
        ' ToBase64CharArray(inArray, offset, length, outArray, outOffset)
        Dim count = Convert.ToBase64CharArray(input, 0, 3, outChars, 0)
        __Check(CStr(count & ":" & New String(outChars, 0, count)), "4:AQID")
    End Sub
End Module
