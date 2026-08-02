' vybe-test: vb/vb_struct_layoutkind/struct_layoutkind
' origin: languages/vb/tests/vb/test_vb_struct_layoutkind.rs

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
Structure UnionType
    <FieldOffset(0)> Public I As Integer
    <FieldOffset(0)> Public B1 As Byte
    <FieldOffset(1)> Public B2 As Byte
    <FieldOffset(2)> Public B3 As Byte
    <FieldOffset(3)> Public B4 As Byte
End Structure

Module M
    Sub Main()
        Dim u As New UnionType()
        u.I = &H12345678
        ' B1 will be the least significant byte on little endian systems (0x78 = 120)
        __Check(CStr(u.B1), "120")
    End Sub
End Module
