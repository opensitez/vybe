use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: GC.AddMemoryPressure, GC.RemoveMemoryPressure & GC Management
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_gc_add_and_remove_memory_pressure() {
    let src = r#"
Imports System

Class NativeBufferHolder
    Private bytesAllocated As Long

    Public Sub New(size As Long)
        bytesAllocated = size
        GC.AddMemoryPressure(bytesAllocated)
    End Sub

    Public Sub Release()
        If bytesAllocated > 0 Then
            GC.RemoveMemoryPressure(bytesAllocated)
            bytesAllocated = 0
        End If
    End Sub
End Class

Module Program
    Sub Main()
        Dim holder As New NativeBufferHolder(1024 * 1024 * 10) ' 10MB pressure
        Console.WriteLine("Added Pressure")
        holder.Release()
        Console.WriteLine("Removed Pressure")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Added Pressure", "Removed Pressure"]);
}

#[test]
fn test_vb_gc_add_memory_pressure_negative_value_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            GC.AddMemoryPressure(-100)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("ArgumentOutOfRangeException Caught on Negative Pressure")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentOutOfRangeException Caught on Negative Pressure"]
    );
}

#[test]
fn test_vb_gc_remove_memory_pressure_negative_value_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            GC.RemoveMemoryPressure(-50)
        Catch ex As ArgumentOutOfRangeException
            Console.WriteLine("ArgumentOutOfRangeException Caught on Negative Remove Pressure")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentOutOfRangeException Caught on Negative Remove Pressure"]
    );
}

#[test]
fn test_vb_gc_suppress_finalize_with_memory_pressure() {
    let src = r#"
Imports System

Class NativeResource
    Implements IDisposable
    Private allocatedBytes As Long

    Public Sub New(bytes As Long)
        allocatedBytes = bytes
        GC.AddMemoryPressure(allocatedBytes)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        If allocatedBytes > 0 Then
            GC.RemoveMemoryPressure(allocatedBytes)
            allocatedBytes = 0
        End If
        GC.SuppressFinalize(Me)
    End Sub

    Protected Overrides Sub Finalize()
        Dispose()
    End Sub
End Class

Module Program
    Sub Main()
        Using res As New NativeResource(5000000)
            Console.WriteLine("Native Resource Active")
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Native Resource Active"]);
}

#[test]
fn test_vb_gc_add_memory_pressure_zero_is_allowed() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        GC.AddMemoryPressure(0)
        GC.RemoveMemoryPressure(0)
        Console.WriteLine("Zero Pressure Passed")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Zero Pressure Passed"]);
}

#[test]
fn test_vb_gc_get_total_allocated_bytes_precise() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim b1 = GC.GetTotalAllocatedBytes(precise:=True)
        Dim dummy As New Byte(1000) {}
        Dim b2 = GC.GetTotalAllocatedBytes(precise:=True)
        Console.WriteLine(b2 > b1)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gc_collection_mode_default_forced_optimized() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        GC.Collect(0, GCCollectionMode.Default)
        GC.Collect(0, GCCollectionMode.Forced)
        GC.Collect(0, GCCollectionMode.Optimized)
        Console.WriteLine("All Collection Modes Succeeded")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["All Collection Modes Succeeded"]);
}

#[test]
fn test_vb_gc_latency_mode_get_and_set() {
    let src = r#"
Imports System
Imports System.Runtime

Module Program
    Sub Main()
        Dim currentMode = GCSettings.LatencyMode
        Console.WriteLine(currentMode.ToString().Length > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gc_is_server_gc_boolean() {
    let src = r#"
Imports System
Imports System.Runtime

Module Program
    Sub Main()
        Dim isServer = GCSettings.IsServerGC
        Console.WriteLine("Server GC Query: " & (isServer OrElse Not isServer))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Server GC Query: True"]);
}

#[test]
fn test_vb_gc_large_object_heap_compaction_mode() {
    let src = r#"
Imports System
Imports System.Runtime

Module Program
    Sub Main()
        GCSettings.LargeObjectHeapCompactionMode = GCLargeObjectHeapCompactionMode.CompactOnce
        Console.WriteLine(GCSettings.LargeObjectHeapCompactionMode.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["CompactOnce"]);
}

#[test]
fn test_vb_gc_get_generation_from_weak_reference() {
    let src = r#"
Imports System

Class Target
End Class

Module Program
    Sub Main()
        Dim obj As New Target()
        Dim weakRef As New WeakReference(obj)
        Dim gen = GC.GetGeneration(weakRef)
        Console.WriteLine(gen >= 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gc_allocate_array_pinned_unpinned() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' GC.AllocateArray(Of T)(length, pinned:=True)
        Dim pinnedArr = GC.AllocateArray(Of Byte)(64, pinned:=True)
        Dim unpinnedArr = GC.AllocateUninitializedArray(Of Integer)(10, pinned:=False)
        Console.WriteLine(pinnedArr.Length & "|" & unpinnedArr.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["64|10"]);
}

#[test]
fn test_vb_gc_allocate_uninitialized_array_primitive() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim ints = GC.AllocateUninitializedArray(Of Integer)(100)
        Console.WriteLine(ints.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_gc_get_configuration_variable() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim info = GC.GetGCMemoryInfo()
        Console.WriteLine(info.TotalAvailableMemoryBytes > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gc_memory_info_heap_count() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim info = GC.GetGCMemoryInfo()
        Console.WriteLine(info.HeapCount >= 1)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gc_memory_info_pause_time_percentage() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim info = GC.GetGCMemoryInfo()
        Console.WriteLine(info.PauseTimePercentage >= 0.0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gc_add_memory_pressure_multiple_allocations() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["Pressure Balanced"]);
}

#[test]
fn test_vb_gc_get_generation_for_null_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            Dim obj As Object = Nothing
            GC.GetGeneration(obj)
        Catch ex As ArgumentNullException
            Console.WriteLine("ArgumentNullException Caught on Null GetGeneration")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException Caught on Null GetGeneration"]
    );
}

#[test]
fn test_vb_gc_force_blocking_full_collection() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        GC.Collect(2, GCCollectionMode.Forced, blocking:=True, compacting:=True)
        Console.WriteLine("Full Compacting GC Completed")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Full Compacting GC Completed"]);
}

#[test]
fn test_vb_gc_get_total_memory_after_gc() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim memBefore = GC.GetTotalMemory(forceFullCollection:=True)
        Console.WriteLine(memBefore > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
