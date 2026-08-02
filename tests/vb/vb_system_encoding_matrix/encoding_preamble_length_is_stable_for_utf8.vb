' vybe-test: vb/vb_system_encoding_matrix/encoding_preamble_length_is_stable_for_utf8
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
        Dim preamble() As Byte = Encoding.UTF8.GetPreamble()
        __Check(CStr(preamble.Length >= 0), "True")
        __Check(CStr(preamble.Length = 3 OrElse preamble.Length = 0), "True")
    End Sub
End Module
