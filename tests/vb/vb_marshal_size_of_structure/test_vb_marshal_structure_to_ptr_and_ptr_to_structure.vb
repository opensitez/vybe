' vybe-test: vb/vb_marshal_size_of_structure/test_vb_marshal_structure_to_ptr_and_ptr_to_structure
' origin: languages/vb/tests/vb/test_vb_marshal_size_of_structure.rs

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

<StructLayout(LayoutKind.Sequential)>
Structure Header
    Public Version As Integer
    Public Flag As Byte
End Structure

Module Program
    Sub Main()
        Dim h As New Header With {.Version = 2, .Flag = 1}
        Dim size = Marshal.SizeOf(GetType(Header))
        Dim ptr As IntPtr = Marshal.AllocHGlobal(size)

        Marshal.StructureToPtr(h, ptr, False)
        Dim restored As Header = CType(Marshal.PtrToStructure(ptr, GetType(Header)), Header)
        Marshal.FreeHGlobal(ptr)

        __Check(CStr(restored.Version & "|" & restored.Flag), "2|1")
    End Sub
End Module
