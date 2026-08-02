' vybe-test: vb/vb_gc_add_memory_pressure/test_vb_gc_add_memory_pressure_multiple_allocations
' origin: languages/vb/tests/vb/test_vb_gc_add_memory_pressure.rs

Imports System

Module Program
    Sub Main()
        For i As Integer = 1 To 5
            GC.AddMemoryPressure(100000)
        Next
        For i As Integer = 1 To 5
            GC.RemoveMemoryPressure(100000)
        Next
        Console.WriteLine("Pressure Balanced")
    End Sub
End Module
