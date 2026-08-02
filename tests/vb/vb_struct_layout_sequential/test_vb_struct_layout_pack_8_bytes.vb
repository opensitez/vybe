' vybe-test: vb/vb_struct_layout_sequential/test_vb_struct_layout_pack_8_bytes
' origin: languages/vb/tests/vb/test_vb_struct_layout_sequential.rs

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

Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, Pack:=8)>
Structure Pack8Struct
    Public A As Byte
    Public B As Double
End Structure

Module Program
    Sub Main()
        ' Offset of B is 8, total size = 16
        __Check(CStr(Marshal.SizeOf(GetType(Pack8Struct))), "16")
    End Sub
End Module
