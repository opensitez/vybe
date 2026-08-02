' vybe-test: vb/vb_system_encoding_matrix/encoding_default_roundtrip
' origin: languages/vb/tests/vb/test_vb_system_encoding_matrix.rs

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

Module M
    Sub Main()
        Dim text As String = "runtime"
        Dim bytes() As Byte = Encoding.Default.GetBytes(text)
        Dim restored As String = Encoding.Default.GetString(bytes)

        __Check(CStr(restored), "runtime")
        __Check(CStr(bytes.Length), "7")
    End Sub
End Module
