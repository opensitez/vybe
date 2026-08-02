' vybe-test: vb/vb_bit_converter_primitive_bytes/test_vb_bit_converter_get_bytes_to_int32
' origin: languages/vb/tests/vb/test_vb_bit_converter_primitive_bytes.rs

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
        Dim val As Integer = 123456789
        Dim bytes As Byte() = BitConverter.GetBytes(val)
        __Check(CStr(bytes.Length), "4")
        Dim restored As Integer = BitConverter.ToInt32(bytes, 0)
        __Check(CStr(restored), "123456789")
    End Sub
End Module
