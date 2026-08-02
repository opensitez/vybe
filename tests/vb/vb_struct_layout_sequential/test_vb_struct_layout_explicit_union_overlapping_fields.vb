' vybe-test: vb/vb_struct_layout_sequential/test_vb_struct_layout_explicit_union_overlapping_fields
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
Structure IntFloatUnion
    <FieldOffset(0)> Public AsInt As Integer
    <FieldOffset(0)> Public AsSingle As Single
End Structure

Module Program
    Sub Main()
        Dim u As New IntFloatUnion With {.AsInt = &H3F800000} ' IEEE 754 float representation of 1.0F
        __Check(CStr(u.AsSingle), "1")
    End Sub
End Module
