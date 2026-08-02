' vybe-test: vb/vb_marshal_size_of_structure/test_vb_marshal_copy_int_array_to_and_from_hglobal
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
        Dim source As Integer() = {10, 20, 30}
        Dim ptr As IntPtr = Marshal.AllocHGlobal(source.Length * 4)
        Marshal.Copy(source, 0, ptr, source.Length)

        Dim dest(2) As Integer
        Marshal.Copy(ptr, dest, 0, dest.Length)
        Marshal.FreeHGlobal(ptr)
        __Check(CStr(String.Join("-", dest)), "10-20-30")
    End Sub
End Module
