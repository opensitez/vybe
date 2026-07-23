use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Standard IDisposable Implementation & Double-Dispose Safety
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_idisposable_double_dispose_idempotent() {
    let src = r#"
Imports System

Class ManagedResource
    Implements IDisposable

    Public Property DisposeCount As Integer = 0
    Private disposedValue As Boolean

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not disposedValue Then
            If disposing Then
                DisposeCount += 1
            End If
            disposedValue = True
        End If
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(disposing:=True)
        GC.SuppressFinalize(Me)
    End Sub
End Class

Module Program
    Sub Main()
        Dim res As New ManagedResource()
        res.Dispose()
        res.Dispose() ' Second call should be no-op!
        Console.WriteLine(res.DisposeCount)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_using_statement_auto_disposes() {
    let src = r#"
Imports System

Class AutoResource
    Implements IDisposable
    Public Property IsDisposed As Boolean = False
    Public Sub Dispose() Implements IDisposable.Dispose
        IsDisposed = True
        Console.WriteLine("AutoResource Disposed")
    End Sub
End Class

Module Program
    Sub Main()
        Using res As New AutoResource()
            Console.WriteLine("Inside Using")
        End Using
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Inside Using", "AutoResource Disposed"]);
}

#[test]
fn test_vb_using_statement_disposes_on_exception() {
    let src = r#"
Imports System

Class ExceptionSafeResource
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Disposed After Exception")
    End Sub
End Class

Module Program
    Sub Main()
        Try
            Using res As New ExceptionSafeResource()
                Throw New InvalidOperationException("Fault inside Using")
            End Using
        Catch ex As InvalidOperationException
            Console.WriteLine("Exception Caught Outside")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Disposed After Exception", "Exception Caught Outside"]
    );
}

#[test]
fn test_vb_idisposable_throw_if_disposed_guard() {
    let src = r#"
Imports System

Class GuardedResource
    Implements IDisposable
    Private isDisposed As Boolean = False

    Public Sub DoWork()
        If isDisposed Then Throw New ObjectDisposedException(NameOf(GuardedResource))
        Console.WriteLine("Work Done")
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        isDisposed = True
    End Sub
End Class

Module Program
    Sub Main()
        Dim res As New GuardedResource()
        res.DoWork()
        res.Dispose()

        Try
            res.DoWork()
        Catch ex As ObjectDisposedException
            Console.WriteLine("ObjectDisposedException Caught")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Work Done", "ObjectDisposedException Caught"]
    );
}

#[test]
fn test_vb_idisposable_derived_class_disposal_pattern() {
    let src = r#"
Imports System

Class BaseRes
    Implements IDisposable
    Protected BaseDisposed As Boolean = False

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not BaseDisposed Then
            If disposing Then Console.WriteLine("Base Managed Cleaned")
            BaseDisposed = True
        End If
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
End Class

Class DerivedRes
    Inherits BaseRes

    Private DerivedDisposed As Boolean = False

    Protected Overrides Sub Dispose(disposing As Boolean)
        If Not DerivedDisposed Then
            If disposing Then Console.WriteLine("Derived Managed Cleaned")
            DerivedDisposed = True
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class

Module Program
    Sub Main()
        Dim res As New DerivedRes()
        res.Dispose()
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Derived Managed Cleaned", "Base Managed Cleaned"]
    );
}

#[test]
fn test_vb_using_statement_multiple_resources_same_type() {
    let src = r#"
Imports System

Class Tracker
    Implements IDisposable
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Disposed " & Name)
    End Sub
End Class

Module Program
    Sub Main()
        Using r1 As New Tracker("R1"), r2 As New Tracker("R2")
            Console.WriteLine("Inside Multi Using")
        End Using
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Inside Multi Using", "Disposed R2", "Disposed R1"]
    );
}

#[test]
fn test_vb_using_statement_null_resource_is_safe_noop() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim res As IDisposable = Nothing
        Using res
            Console.WriteLine("Inside Null Using")
        End Using
        Console.WriteLine("After Null Using")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Inside Null Using", "After Null Using"]);
}

#[test]
fn test_vb_using_statement_struct_disposable() {
    let src = r#"
Imports System

Structure StructDisposable
    Implements IDisposable

    Public Property ID As Integer
    Public Sub New(i As Integer)
        ID = i
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Struct Disposed " & ID)
    End Sub
End Structure

Module Program
    Sub Main()
        Using s As New StructDisposable(42)
            Console.WriteLine("Inside Struct Using")
        End Using
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Inside Struct Using", "Struct Disposed 42"]
    );
}

#[test]
fn test_vb_iasyncdisposable_dispose_async_pattern() {
    let src = r#"
Imports System
Imports System.Threading.Tasks

Class AsyncResource
    Implements IAsyncDisposable

    Public Async Function DisposeAsync() As ValueTask Implements IAsyncDisposable.DisposeAsync
        Await Task.Yield()
        Console.WriteLine("Async Disposed")
    End Function
End Class

Module Program
    Sub Main()
        Dim res As New AsyncResource()
        Dim vt = res.DisposeAsync()
        vt.AsTask().Wait()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Async Disposed"]);
}

#[test]
fn test_vb_composite_disposable_disposes_children() {
    let src = r#"
Imports System
Imports System.Collections.Generic

Class CompositeDisposable
    Implements IDisposable
    Private children As New List(Of IDisposable)()

    Public Sub Add(item As IDisposable)
        children.Add(item)
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        For Each child In children
            child.Dispose()
        Next
        children.Clear()
    End Sub
End Class

Class ChildRes
    Implements IDisposable
    Private tag As String
    Public Sub New(t As String)
        tag = t
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Disposed Child " & tag)
    End Sub
End Class

Module Program
    Sub Main()
        Dim comp As New CompositeDisposable()
        comp.Add(New ChildRes("A"))
        comp.Add(New ChildRes("B"))
        comp.Dispose()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Disposed Child A", "Disposed Child B"]);
}

#[test]
fn test_vb_idisposable_explicit_interface_implementation() {
    let src = r#"
Imports System

Class ExplicitDisposable
    Implements IDisposable

    Private Sub IDisposable_Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Explicit IDisposable.Dispose")
    End Sub
End Class

Module Program
    Sub Main()
        Dim ed As New ExplicitDisposable()
        ' Must cast to IDisposable to call Dispose!
        Dim d As IDisposable = ed
        d.Dispose()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Explicit IDisposable.Dispose"]);
}

#[test]
fn test_vb_using_declaration_statement_scoping() {
    let src = r#"
Imports System

Class ScopeTracker
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("ScopeTracker Disposed")
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Using res As New ScopeTracker()
                Console.WriteLine("Doing Work in Inner Scope")
            End Using
        End Sub()
        Console.WriteLine("Outer Scope")
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec![
            "Doing Work in Inner Scope",
            "ScopeTracker Disposed",
            "Outer Scope"
        ]
    );
}

#[test]
fn test_vb_idisposable_exception_in_dispose_handled() {
    let src = r#"
Imports System

Class ThrowingDisposable
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Throw New InvalidOperationException("Dispose Failed")
    End Sub
End Class

Module Program
    Sub Main()
        Try
            Using res As New ThrowingDisposable()
                Console.WriteLine("Inside Using Block")
            End Using
        Catch ex As InvalidOperationException
            Console.WriteLine(ex.Message)
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Inside Using Block", "Dispose Failed"]);
}

#[test]
fn test_vb_idisposable_resource_reinitialization_throws() {
    let src = r#"
Imports System

Class SingleUseResource
    Implements IDisposable
    Private isDisposed As Boolean = False

    Public Sub Initialize()
        If isDisposed Then Throw New ObjectDisposedException("SingleUseResource")
        Console.WriteLine("Initialized")
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        isDisposed = True
    End Sub
End Class

Module Program
    Sub Main()
        Dim res As New SingleUseResource()
        res.Initialize()
        res.Dispose()
        Try
            res.Initialize()
        Catch ex As ObjectDisposedException
            Console.WriteLine("Cannot Reinitialize Disposed Resource")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Initialized", "Cannot Reinitialize Disposed Resource"]
    );
}

#[test]
fn test_vb_using_statement_with_expression_target() {
    let src = r#"
Imports System

Class Factory
    Public Shared Function Create() As FactoryResource
        Return New FactoryResource()
    End Function
End Class

Class FactoryResource
    Implements IDisposable
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Factory Resource Disposed")
    End Sub
End Class

Module Program
    Sub Main()
        Using res = Factory.Create()
            Console.WriteLine("Using Factory Result")
        End Using
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Using Factory Result", "Factory Resource Disposed"]
    );
}

#[test]
fn test_vb_idisposable_thread_safe_double_dispose() {
    let src = r#"
Imports System
Imports System.Threading

Class ThreadSafeDisposable
    Implements IDisposable
    Private disposeState As Integer = 0

    Public Sub Dispose() Implements IDisposable.Dispose
        ' Ensure only one thread executes disposal logic!
        If Interlocked.Exchange(disposeState, 1) = 0 Then
            Console.WriteLine("Thread-Safe Disposal Executed")
        End If
    End Sub
End Class

Module Program
    Sub Main()
        Dim tsd As New ThreadSafeDisposable()
        tsd.Dispose()
        tsd.Dispose()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Thread-Safe Disposal Executed"]);
}

#[test]
fn test_vb_idisposable_interop_stream_wrapper() {
    let src = r#"
Imports System
Imports System.IO

Module Program
    Sub Main()
        Dim ms As New MemoryStream()
        ms.WriteByte(65)
        ms.Dispose()

        Try
            ms.WriteByte(66)
        Catch ex As ObjectDisposedException
            Console.WriteLine("MemoryStream Disposed Safely")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["MemoryStream Disposed Safely"]);
}

#[test]
fn test_vb_using_statement_reassigning_variable_disposes_original() {
    let src = r#"
Imports System

Class NamedRes
    Implements IDisposable
    Public Name As String
    Public Sub New(n As String)
        Name = n
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Disposed " & Name)
    End Sub
End Class

Module Program
    Sub Main()
        Dim r As New NamedRes("Original")
        Using r
            Console.WriteLine("Inside Using Original")
        End Using
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Inside Using Original", "Disposed Original"]
    );
}

#[test]
fn test_vb_idisposable_nested_using_order() {
    let src = r#"
Imports System

Class LevelTracker
    Implements IDisposable
    Private level As String
    Public Sub New(l As String)
        level = l
    End Sub
    Public Sub Dispose() Implements IDisposable.Dispose
        Console.WriteLine("Exit Level " & level)
    End Sub
End Class

Module Program
    Sub Main()
        Using Outer As New LevelTracker("Outer")
            Using Inner As New LevelTracker("Inner")
                Console.WriteLine("Innermost Action")
            End Using
        End Using
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Innermost Action", "Exit Level Inner", "Exit Level Outer"]
    );
}

#[test]
fn test_vb_idisposable_check_disposed_property() {
    let src = r#"
Imports System

Class ReadableStateDisposable
    Implements IDisposable

    Public ReadOnly Property IsDisposed As Boolean

    Public Sub Dispose() Implements IDisposable.Dispose
        _IsDisposed = True
        Console.WriteLine("Disposed State Flag Set")
    End Sub
End Class

Module Program
    Sub Main()
        Dim r As New ReadableStateDisposable()
        Console.WriteLine(r.IsDisposed)
        r.Dispose()
        Console.WriteLine(r.IsDisposed)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["False", "Disposed State Flag Set", "True"]
    );
}
