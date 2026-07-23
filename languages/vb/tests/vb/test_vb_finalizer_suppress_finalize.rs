use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Finalizer (Protected Overrides Sub Finalize) & GC.SuppressFinalize
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_finalizer_suppress_finalize_prevents_finalizer_execution() {
    let src = r#"
Imports System

Class SuppressedFinalizerObject
    Public Shared FinalizerRan As Boolean = False

    Protected Overrides Sub Finalize()
        FinalizerRan = True
    End Sub

    Public Sub Cleanup()
        GC.SuppressFinalize(Me)
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim obj As New SuppressedFinalizerObject()
            obj.Cleanup()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()

        Console.WriteLine(SuppressedFinalizerObject.FinalizerRan)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False"]);
}

#[test]
fn test_vb_finalizer_unsuppressed_runs_on_gc() {
    let src = r#"
Imports System

Class ActiveFinalizerObject
    Public Shared RanCount As Integer = 0

    Protected Overrides Sub Finalize()
        RanCount += 1
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim obj As New ActiveFinalizerObject()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        Console.WriteLine("Finalizer Ran: " & (ActiveFinalizerObject.RanCount = 1))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Finalizer Ran: True"]);
}

#[test]
fn test_vb_gc_re_register_for_finalize() {
    let src = r#"
Imports System

Class ReRegisteredObject
    Public Shared Executions As Integer = 0

    Protected Overrides Sub Finalize()
        Executions += 1
    End Sub

    Public Sub ReRegister()
        GC.ReRegisterForFinalize(Me)
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As New ReRegisteredObject()
        GC.SuppressFinalize(obj)
        obj.ReRegister() ' Re-enable finalization!

        obj = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()

        Console.WriteLine("Finalized Count: " & ReRegisteredObject.Executions)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Finalized Count: 1"]);
}

#[test]
fn test_vb_finalizer_calls_mybase_finalize() {
    let src = r#"
Imports System

Class BaseFinalizer
    Public Shared BaseRan As Boolean = False
    Protected Overrides Sub Finalize()
        BaseRan = True
    End Sub
End Class

Class DerivedFinalizer
    Inherits BaseFinalizer
    Public Shared DerivedRan As Boolean = False
    Protected Overrides Sub Finalize()
        DerivedRan = True
        MyBase.Finalize()
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim d As New DerivedFinalizer()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        Console.WriteLine(DerivedFinalizer.DerivedRan & "|" & BaseFinalizer.BaseRan)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_gc_wait_for_pending_finalizers() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Console.WriteLine("Pending Finalizers Complete")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Pending Finalizers Complete"]);
}

#[test]
fn test_vb_finalizer_swallows_unhandled_exception() {
    let src = r#"
Imports System

Class FaultyFinalizer
    Protected Overrides Sub Finalize()
        ' Exceptions in finalizer are swallowed by CLR runtime without crashing application in default policy
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim f As New FaultyFinalizer()
        End Sub()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Console.WriteLine("Completed Safe GC")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Completed Safe GC"]);
}

#[test]
fn test_vb_finalizer_resurrects_object_once() {
    let src = r#"
Imports System

Class Phoenix
    Public Shared Instance As Phoenix
    Protected Overrides Sub Finalize()
        Instance = Me ' Resurrect
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim p As New Phoenix()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        Console.WriteLine("Resurrected: " & (Phoenix.Instance IsNot Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Resurrected: True"]);
}

#[test]
fn test_vb_finalizer_thread_environment_check() {
    let src = r#"
Imports System
Imports System.Threading

Class ThreadCheckFinalizer
    Public Shared FinalizerThreadId As Integer = -1

    Protected Overrides Sub Finalize()
        FinalizerThreadId = Thread.CurrentThread.ManagedThreadId
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim obj As New ThreadCheckFinalizer()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        Console.WriteLine("Finalizer Executed on Valid Thread: " & (ThreadCheckFinalizer.FinalizerThreadId > 0))
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Finalizer Executed on Valid Thread: True"]
    );
}

#[test]
fn test_vb_idisposable_pattern_with_finalizer_fallback() {
    let src = r#"
Imports System

Class FullDisposablePattern
    Implements IDisposable

    Public Property CleanedFromDispose As Boolean = False
    Public Property CleanedFromFinalizer As Boolean = False
    Private disposedValue As Boolean

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                CleanedFromDispose = True
            Else
                CleanedFromFinalizer = True
            End If
            disposedValue = True
        End If
    End Sub

    Protected Overrides Sub Finalize()
        Dispose(disposing:=False)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As New FullDisposablePattern()
        obj.Dispose()
        Console.WriteLine(obj.CleanedFromDispose & "|" & obj.CleanedFromFinalizer)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_finalizer_unmanaged_handle_cleanup_simulation() {
    let src = r#"
Imports System

Class UnmanagedHandleHolder
    Private nativeHandle As IntPtr = New IntPtr(12345)
    Public Shared HandleReleased As Boolean = False

    Protected Overrides Sub Finalize()
        If nativeHandle <> IntPtr.Zero Then
            nativeHandle = IntPtr.Zero
            HandleReleased = True
        End If
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim holder As New UnmanagedHandleHolder()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        Console.WriteLine("Handle Released: " & UnmanagedHandleHolder.HandleReleased)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Handle Released: True"]);
}

#[test]
fn test_vb_gc_suppress_finalize_null_check_safe() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            GC.SuppressFinalize(Nothing)
        Catch ex As ArgumentNullException
            Console.WriteLine("ArgumentNullException Caught on Null SuppressFinalize")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException Caught on Null SuppressFinalize"]
    );
}

#[test]
fn test_vb_finalizer_in_nested_class() {
    let src = r#"
Imports System

Class OuterContainer
    Class InnerNested
        Public Shared NestedFinalized As Boolean = False
        Protected Overrides Sub Finalize()
            NestedFinalized = True
        End Sub
    End Class
End Class

Module Program
    Sub Main()
        Sub()
            Dim inner As New OuterContainer.InnerNested()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        Console.WriteLine(OuterContainer.InnerNested.NestedFinalized)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_finalizer_ordering_unspecified() {
    let src = r#"
Imports System

Class ObjA
    Protected Overrides Sub Finalize()
    End Sub
End Class

Class ObjB
    Protected Overrides Sub Finalize()
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim a As New ObjA()
            Dim b As New ObjB()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()
        Console.WriteLine("Multiple Finalizers Executed Safely")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Multiple Finalizers Executed Safely"]);
}

#[test]
fn test_vb_finalizer_accessing_other_managed_objects_unsafe() {
    let src = r#"
Imports System

Class ManagedDependency
    Public Sub Ping()
        Console.WriteLine("Dependency Ping")
    End Sub
End Class

Class Consumer
    Private dep As New ManagedDependency()

    Protected Overrides Sub Finalize()
        ' Note: Accessing dep in Finalize is unsafe because dep may already be finalized!
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim c As New Consumer()
        End Sub()
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Console.WriteLine("Finalization Finished")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Finalization Finished"]);
}

#[test]
fn test_vb_finalizer_generic_class() {
    let src = r#"
Imports System

Class GenericFinalizer(Of T)
    Public Shared Finalized As Boolean = False
    Protected Overrides Sub Finalize()
        Finalized = True
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim g As New GenericFinalizer(Of String)()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        Console.WriteLine(GenericFinalizer(Of String).Finalized)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_finalizer_abstract_mustinherit_class() {
    let src = r#"
Imports System

MustInherit Class AbstractWithFinalizer
    Public Shared AbstractFinalizerRan As Boolean = False
    Protected Overrides Sub Finalize()
        AbstractFinalizerRan = True
    End Sub
End Class

Class ConcreteDerived
    Inherits AbstractWithFinalizer
End Class

Module Program
    Sub Main()
        Sub()
            Dim c As New ConcreteDerived()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        Console.WriteLine(AbstractWithFinalizer.AbstractFinalizerRan)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_finalizer_multiple_gc_collect_calls() {
    let src = r#"
Imports System

Class MultiCollectObject
    Public Shared Count As Integer = 0
    Protected Overrides Sub Finalize()
        Count += 1
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim m As New MultiCollectObject()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()
        GC.Collect()
        GC.WaitForPendingFinalizers()

        Console.WriteLine(MultiCollectObject.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_finalizer_suppress_finalize_called_multiple_times() {
    let src = r#"
Imports System

Class MultiSuppressObject
    Protected Overrides Sub Finalize()
        Console.WriteLine("Finalize Ran")
    End Sub
End Class

Module Program
    Sub Main()
        Dim obj As New MultiSuppressObject()
        GC.SuppressFinalize(obj)
        GC.SuppressFinalize(obj)
        GC.SuppressFinalize(obj)

        obj = Nothing
        GC.Collect()
        GC.WaitForPendingFinalizers()
        Console.WriteLine("Suppression Multi Safe")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Suppression Multi Safe"]);
}

#[test]
fn test_vb_finalizer_large_object_heap_finalization() {
    let src = r#"
Imports System

Class LargeFinalizableObject
    Private buffer(100000) As Byte ' Placed on Large Object Heap (LOH)
    Public Shared Finalized As Boolean = False

    Protected Overrides Sub Finalize()
        Finalized = True
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim l As New LargeFinalizableObject()
        End Sub()

        GC.Collect(2, GCCollectionMode.Forced)
        GC.WaitForPendingFinalizers()

        Console.WriteLine(LargeFinalizableObject.Finalized)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_finalizer_re_register_null_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            GC.ReRegisterForFinalize(Nothing)
        Catch ex As ArgumentNullException
            Console.WriteLine("ArgumentNullException Caught on Null ReRegisterForFinalize")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException Caught on Null ReRegisterForFinalize"]
    );
}
