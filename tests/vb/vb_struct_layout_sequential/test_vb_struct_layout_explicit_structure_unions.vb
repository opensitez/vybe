' vybe-test: vb/vb_struct_layout_sequential/test_vb_struct_layout_explicit_structure_unions
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

<StructLayout(LayoutKind.Explicit)>
Structure ColorPixel
    <FieldOffset(0)> Public R As Byte
    <FieldOffset(1)> Public G As Byte
    <FieldOffset(2)> Public B As Byte
    <FieldOffset(3)> Public A As Byte
    <FieldOffset(0)> Public RgbaValue As UInteger
End Structure

Module Program
    Sub Main()
        Dim p As New ColorPixel With {.RgbaValue = &HFF0000FFUI}
        __Check(CStr(Marshal.SizeOf(GetType(ColorPixel)) & "|" & p.RgbaValue <> 0), "4|True")
    End Sub
End Module
