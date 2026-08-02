' vybe-test: vb/vb_struct_layout_sequential/test_vb_struct_layout_offsets_verification
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

<StructLayout(LayoutKind.Sequential, Pack:=4)>
Structure MixedStruct
    Public B1 As Byte
    Public I1 As Integer
    Public B2 As Byte
    Public L1 As Long
End Structure

Module Program
    Sub Main()
        Dim offB1 = Marshal.OffsetOf(GetType(MixedStruct), "B1").ToInt32()
        Dim offI1 = Marshal.OffsetOf(GetType(MixedStruct), "I1").ToInt32()
        Dim offB2 = Marshal.OffsetOf(GetType(MixedStruct), "B2").ToInt32()
        Dim offL1 = Marshal.OffsetOf(GetType(MixedStruct), "L1").ToInt32()
        __Check(CStr(offB1 & "|" & offI1 & "|" & offB2 & "|" & offL1), "0|4|8|12")
    End Sub
End Module
