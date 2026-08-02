' vybe-test: vb/vb_system_encoding_matrix/encoding_max_byte_count_is_non_negative_for_ascii
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
        __Check(CStr(Encoding.UTF8.GetMaxByteCount(0) >= 0), "True")
        __Check(CStr(Encoding.Unicode.GetMaxByteCount(1)), "2")
    End Sub
End Module
