' vybe-test: vb/vb_marshal_size_of_structure/test_vb_marshal_copy_byte_array_to_and_from_hglobal
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

Module Program
    Sub Main()
        Dim source As Byte() = {100, 101, 102, 103}
        Dim ptr As IntPtr = Marshal.AllocHGlobal(source.Length)
        Marshal.Copy(source, 0, ptr, source.Length)

        Dim dest(3) As Byte
        Marshal.Copy(ptr, dest, 0, dest.Length)
        Marshal.FreeHGlobal(ptr)
        __Check(CStr(String.Join(",", dest)), "100,101,102,103")
    End Sub
End Module
