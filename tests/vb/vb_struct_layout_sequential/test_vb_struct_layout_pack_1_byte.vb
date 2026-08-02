' vybe-test: vb/vb_struct_layout_sequential/test_vb_struct_layout_pack_1_byte
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

<StructLayout(LayoutKind.Sequential, Pack:=1)>
Structure PackedStruct
    Public A As Byte
    Public B As Integer
End Structure

Module Program
    Sub Main()
        __Check(CStr(Marshal.SizeOf(GetType(PackedStruct))), "5")
    End Sub
End Module
