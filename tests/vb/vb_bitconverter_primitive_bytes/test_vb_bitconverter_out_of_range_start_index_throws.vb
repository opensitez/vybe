' vybe-test: vb/vb_bitconverter_primitive_bytes/test_vb_bitconverter_out_of_range_start_index_throws
' origin: languages/vb/tests/vb/test_vb_bitconverter_primitive_bytes.rs

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
        Dim bytes As Byte() = {1, 2, 3}
        Try
            BitConverter.ToInt32(bytes, 1) ' Needs 4 bytes, only 2 left!
        Catch ex As ArgumentException
            __Check(CStr("ArgumentException Caught on Truncated Buffer"), "ArgumentException Caught on Truncated Buffer")
        End Try
    End Sub
End Module
