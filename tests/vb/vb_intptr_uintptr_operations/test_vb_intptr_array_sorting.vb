' vybe-test: vb/vb_intptr_uintptr_operations/test_vb_intptr_array_sorting
' origin: languages/vb/tests/vb/test_vb_intptr_uintptr_operations.rs

Imports System

Module Program
    Sub Main()
        Dim pointers As IntPtr() = {New IntPtr(30), New IntPtr(10), New IntPtr(20)}
        Array.Sort(pointers)
        For Each p In pointers
            Console.WriteLine(p.ToInt32())
        Next
    End Sub
End Module
