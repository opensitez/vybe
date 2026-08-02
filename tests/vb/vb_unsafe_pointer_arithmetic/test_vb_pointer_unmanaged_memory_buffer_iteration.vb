' vybe-test: vb/vb_unsafe_pointer_arithmetic/test_vb_pointer_unmanaged_memory_buffer_iteration
' origin: languages/vb/tests/vb/test_vb_unsafe_pointer_arithmetic.rs

Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim count = 5
        Dim ptr As IntPtr = Marshal.AllocHGlobal(count * 4)

        ' Write integers 1 to 5
        For i As Integer = 0 To count - 1
            Marshal.WriteInt32(IntPtr.Add(ptr, i * 4), (i + 1) * 10)
        Next

        ' Read back
        Dim results As New System.Collections.Generic.List(Of String)()
        For i As Integer = 0 To count - 1
            results.Add(Marshal.ReadInt32(IntPtr.Add(ptr, i * 4)).ToString())
        Next
        Marshal.FreeHGlobal(ptr)

        Console.WriteLine(String.Join(",", results))
    End Sub
End Module
