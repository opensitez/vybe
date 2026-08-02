' vybe-test: vb/vb_struct_layout_sequential/test_vb_struct_layout_copy_structure_to_ptr_with_byval_string
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

Imports System
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
Structure PersonNative
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=8)>
    Public Name As String
End Structure

Module Program
    Sub Main()
        Dim p As New PersonNative With {.Name = "Vybe"}
        Dim ptr = Marshal.AllocHGlobal(8)
        Marshal.StructureToPtr(p, ptr, False)

        Dim restored As PersonNative = CType(Marshal.PtrToStructure(ptr, GetType(PersonNative)), PersonNative)
        Marshal.FreeHGlobal(ptr)
        __Check(CStr(restored.Name), "Vybe")
    End Sub
End Module
