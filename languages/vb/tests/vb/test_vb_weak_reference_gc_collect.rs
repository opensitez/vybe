use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: WeakReference(Of T) & GC.Collect Lifetime Tracking
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_weak_reference_generic_target_alive() {
    let src = r#"
Imports System

Class Payload
    Public Property Data As String = "AliveData"
End Class

Module Program
    Sub Main()
        Dim obj As New Payload()
        Dim weakRef As New WeakReference(Of Payload)(obj)

        Dim target As Payload = Nothing
        Dim isAlive = weakRef.TryGetTarget(target)
        Console.WriteLine(isAlive & "|" & target.Data)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|AliveData"]);
}

#[test]
fn test_vb_weak_reference_gc_collect_clears_target() {
    let src = r#"
Imports System

Class DisposableTarget
End Class

Module Program
    Sub Main()
        Dim weakRef As WeakReference(Of DisposableTarget)
        Sub()
            Dim obj As New DisposableTarget()
            weakRef = New WeakReference(Of DisposableTarget)(obj)
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()

        Dim target As DisposableTarget = Nothing
        Dim isAlive = weakRef.TryGetTarget(target)
        Console.WriteLine(isAlive)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_weak_reference_non_generic_target_property() {
    let src = r#"
Imports System

Class Sample
    Public Val As Integer = 42
End Class

Module Program
    Sub Main()
        Dim obj As New Sample()
        Dim weakRef As New WeakReference(obj)
        Dim target As Sample = CType(weakRef.Target, Sample)
        Console.WriteLine(weakRef.IsAlive & "|" & target.Val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|42"]);
}

#[test]
fn test_vb_weak_reference_track_resurrection() {
    let src = r#"
Imports System

Class ResurrectedObject
    Public Shared Holder As ResurrectedObject
    Protected Overrides Sub Finalize()
        Holder = Me ' Resurrect object!
    End Sub
End Class

Module Program
    Sub Main()
        Dim weakRefShort As WeakReference
        Dim weakRefLong As WeakReference

        Sub()
            Dim obj As New ResurrectedObject()
            weakRefShort = New WeakReference(obj, trackResurrection:=False)
            weakRefLong = New WeakReference(obj, trackResurrection:=True)
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        Console.WriteLine("LongTrackAlive: " & (ResurrectedObject.Holder IsNot Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["LongTrackAlive: True"]);
}

#[test]
fn test_vb_gc_get_generation_tracking() {
    let src = r#"
Imports System

Class PersistentObj
End Class

Module Program
    Sub Main()
        Dim obj As New PersistentObj()
        Dim gen0 = GC.GetGeneration(obj)
        Console.WriteLine(gen0 >= 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gc_collection_count_increments() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim initialGen0Count = GC.CollectionCount(0)
        GC.Collect(0)
        Dim newGen0Count = GC.CollectionCount(0)
        Console.WriteLine(newGen0Count > initialGen0Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_weak_reference_cache_simulation() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Class CacheManager
    Private cache As New Dictionary(Of String, WeakReference(Of String))()

    Public Sub Add(key As String, val As String)
        cache(key) = New WeakReference(Of String)(val)
    End Sub

    Public Function GetVal(key As String) As String
        Dim weakRef As WeakReference(Of String) = Nothing
        If cache.TryGetValue(key, weakRef) Then
            Dim target As String = Nothing
            If weakRef.TryGetTarget(target) Then Return target
        End If
        Return Nothing
    End Function
End Class

Module Program
    Sub Main()
        Dim cm As New CacheManager()
        Dim item = "CachedValue"
        cm.Add("K1", item)
        Console.WriteLine(cm.GetVal("K1"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["CachedValue"]);
}

#[test]
fn test_vb_gc_get_total_memory_allocated() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytesAllocated = GC.GetTotalMemory(forceFullCollection:=False)
        Console.WriteLine(bytesAllocated > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gc_max_generation_constant() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' MaxGeneration is typically 2 in standard CLR
        Console.WriteLine(GC.MaxGeneration >= 2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_weak_reference_set_target_reassignment() {
    let src = r#"
Imports System

Class Token
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
End Class

Module Program
    Sub Main()
        Dim t1 As New Token("T1")
        Dim t2 As New Token("T2")
        Dim weakRef As New WeakReference(Of Token)(t1)

        weakRef.SetTarget(t2)
        Dim target As Token = Nothing
        weakRef.TryGetTarget(target)
        Console.WriteLine(target.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["T2"]);
}

#[test]
fn test_vb_gc_register_for_full_gc_notification() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            GC.RegisterForFullGCNotification(10, 10)
            GC.CancelFullGCNotification()
            Console.WriteLine("GC Notification Handled")
        Catch ex As Exception
            Console.WriteLine("GC Notification Unsupported or Ok")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["GC Notification Handled"]);
}

#[test]
fn test_vb_weak_reference_null_target_initialization() {
    let src = r#"
Imports System

Class NullableTarget
End Class

Module Program
    Sub Main()
        Dim weakRef As New WeakReference(Of NullableTarget)(Nothing)
        Dim target As NullableTarget = Nothing
        Dim ok = weakRef.TryGetTarget(target)
        Console.WriteLine(ok & "|" & (target Is Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False|True"]);
}

#[test]
fn test_vb_gc_keep_alive_prevents_premature_collection() {
    let src = r#"
Imports System

Class ResourceTracker
    Public Id As Integer = 100
End Class

Module Program
    Sub Main()
        Dim res As New ResourceTracker()
        Dim id = res.Id
        GC.KeepAlive(res) ' Ensures res is not collected prior to this line!
        Console.WriteLine(id)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100"]);
}

#[test]
fn test_vb_gc_collect_specific_generation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' Force collection of generation 0 and 1
        GC.Collect(1, GCCollectionMode.Forced)
        Console.WriteLine("Gen 1 Collected")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Gen 1 Collected"]);
}

#[test]
fn test_vb_weak_reference_array_of_weak_references() {
    let src = r#"
Imports System

Class Item
    Public Tag As String
    Public Sub New(t As String)
        Tag = t
    End Sub
End Class

Module Program
    Sub Main()
        Dim i1 As New Item("A")
        Dim i2 As New Item("B")
        Dim refs As WeakReference(Of Item)() = {New WeakReference(Of Item)(i1), New WeakReference(Of Item)(i2)}

        For Each r In refs
            Dim item As Item = Nothing
            r.TryGetTarget(item)
            Console.WriteLine(item.Tag)
        Next
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A", "B"]);
}

#[test]
fn test_vb_gc_total_pause_duration_property() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' TotalPauseDuration in modern .NET
        Dim pause = GC.GetTotalPauseDuration()
        Console.WriteLine(pause.TotalMilliseconds >= 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_weak_reference_value_type_boxing() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        ' Boxed value type target in WeakReference
        Dim boxed As Object = 999
        Dim weakRef As New WeakReference(boxed)
        Console.WriteLine(weakRef.Target.ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["999"]);
}

#[test]
fn test_vb_gc_get_allocated_bytes_for_current_thread() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes = GC.GetAllocatedBytesForCurrentThread()
        Console.WriteLine(bytes > 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_weak_reference_struct_wrapper() {
    let src = r#"
Imports System

Structure ManagedWeakHandle(Of T As Class)
    Private ref As WeakReference(Of T)
    Public Sub New(target As T)
        ref = New WeakReference(Of T)(target)
    End Sub
    Public ReadOnly Property Value As T
        Get
            Dim target As T = Nothing
            If ref IsNot Nothing AndAlso ref.TryGetTarget(target) Then Return target
            Return Nothing
        End Get
    End Property
End Structure

Class Node
    Public Name As String = "StructWrappedNode"
End Class

Module Program
    Sub Main()
        Dim n As New Node()
        Dim handle As New ManagedWeakHandle(Of Node)(n)
        Console.WriteLine(handle.Value.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["StructWrappedNode"]);
}

#[test]
fn test_vb_gc_collection_mode_optimized_check() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        GC.Collect(2, GCCollectionMode.Optimized)
        Console.WriteLine("Optimized Collection Done")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Optimized Collection Done"]);
}
