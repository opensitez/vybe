' vybe-test: vb/vb_system_bitconverter_matrix/bitconverter_tohex_preserves_byte_count
' origin: languages/vb/tests/vb/test_vb_system_bitconverter_matrix.rs

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
        Dim bytes() As Byte = {&H01, &H2A, &HFF}
        Dim text As String = BitConverter.ToString(bytes)
        __Check(CStr(text.Length), "8")
        __Check(CStr(text.StartsWith("01-2A")), "True")
    End Sub
End Module
