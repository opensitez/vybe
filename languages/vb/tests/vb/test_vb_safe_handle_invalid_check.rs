use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Runtime.InteropServices.SafeHandle & Custom Handles
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_custom_safe_handle_is_invalid_check() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class DummySafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Console.WriteLine("Handle Released")
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New DummySafeHandle()
        Console.WriteLine(h.IsInvalid & "|" & h.IsClosed)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_custom_safe_handle_valid_release() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class ValidSafeHandle
    Inherits SafeHandle

    Public Sub New(validPtr As IntPtr)
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(validPtr)
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero OrElse handle = New IntPtr(-1)
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Console.WriteLine("ReleaseHandle Called for Ptr: " & handle.ToInt64())
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New ValidSafeHandle(New IntPtr(999))
        Console.WriteLine("IsInvalid: " & h.IsInvalid)
        h.Dispose()
        Console.WriteLine("IsClosed: " & h.IsClosed)
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec![
            "IsInvalid: False",
            "ReleaseHandle Called for Ptr: 999",
            "IsClosed: True"
        ]
    );
}

#[test]
fn test_vb_safe_handle_dangerous_add_ref_and_release() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class RefCountSafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(New IntPtr(555))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New RefCountSafeHandle()
        Dim success = False
        h.DangerousAddRef(success)
        Console.WriteLine("Ref Added: " & success)
        If success Then h.DangerousRelease()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Ref Added: True"]);
}

#[test]
fn test_vb_safe_handle_dangerous_get_handle_ptr() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class PtrSafeHandle
    Inherits SafeHandle

    Public Sub New(ptr As IntPtr)
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(ptr)
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim rawPtr As New IntPtr(8888)
        Dim h As New PtrSafeHandle(rawPtr)
        Dim retrievedPtr = h.DangerousGetHandle()
        Console.WriteLine(retrievedPtr = rawPtr)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_safe_handle_set_handle_as_invalid() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class ResettableSafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(New IntPtr(100))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Public Sub Invalidate()
        SetHandleAsInvalid()
    End Sub

    Protected Overrides Function ReleaseHandle() As Boolean
        Console.WriteLine("Released")
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New ResettableSafeHandle()
        h.Invalidate()
        Console.WriteLine(h.IsClosed & "|" & h.IsInvalid)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_safe_handle_zero_or_minus_one_is_invalid() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class StandardSafeHandle
    Inherits SafeHandleZeroOrMinusOneIsInvalid

    Public Sub New(ownsHandle As Boolean)
        MyBase.New(ownsHandle)
        SetHandle(New IntPtr(-1))
    End Sub

    Protected Overrides Function ReleaseHandle() As Boolean
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New StandardSafeHandle(True)
        ' SafeHandleZeroOrMinusOneIsInvalid considers -1 as invalid!
        Console.WriteLine(h.IsInvalid)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_safe_handle_critical_finalizer_object_heritage() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class CriticalSafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New CriticalSafeHandle()
        ' Inherits CriticalFinalizerObject implicitly
        Console.WriteLine(TypeOf h Is ConstrainedExecution.CriticalFinalizerObject)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_safe_handle_close_method_calls_dispose() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class ClosableSafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(New IntPtr(123))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Console.WriteLine("ReleaseHandle Executed")
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New ClosableSafeHandle()
        h.Close()
        Console.WriteLine(h.IsClosed)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ReleaseHandle Executed", "True"]);
}

#[test]
fn test_vb_safe_handle_double_dispose_safe_once_release() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class OnceReleaseSafeHandle
    Inherits SafeHandle
    Public Shared ReleaseCount As Integer = 0

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(New IntPtr(777))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        ReleaseCount += 1
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New OnceReleaseSafeHandle()
        h.Dispose()
        h.Dispose() ' Second call ignored!
        Console.WriteLine(OnceReleaseSafeHandle.ReleaseCount)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_safe_handle_owns_handle_false_does_not_call_release() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class NonOwnerSafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=False) ' Does not own handle!
        SetHandle(New IntPtr(444))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Console.WriteLine("Should Not Be Called")
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New NonOwnerSafeHandle()
        h.Dispose()
        Console.WriteLine("NonOwner Disposed Safely")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["NonOwner Disposed Safely"]);
}

#[test]
fn test_vb_safe_handle_in_using_block() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class UsingSafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(New IntPtr(10))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Console.WriteLine("Using SafeHandle Released")
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Using h As New UsingSafeHandle()
            Console.WriteLine("Inside Using SafeHandle")
        End Using
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["Inside Using SafeHandle", "Using SafeHandle Released"]
    );
}

#[test]
fn test_vb_safe_handle_cannot_use_after_closed_throws() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class StrictSafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(New IntPtr(1))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New StrictSafeHandle()
        h.Dispose()

        Try
            Dim success = False
            h.DangerousAddRef(success)
        Catch ex As ObjectDisposedException
            Console.WriteLine("ObjectDisposedException Caught on Closed SafeHandle")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ObjectDisposedException Caught on Closed SafeHandle"]
    );
}

#[test]
fn test_vb_safe_handle_gchandle_interop_combination() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim data = "InteropPayload"
        Dim gcHandle = GCHandle.Alloc(data, GCHandleType.Pinned)
        Dim ptr = gcHandle.AddrOfPinnedObject()
        Console.WriteLine(ptr <> IntPtr.Zero)
        gcHandle.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_safe_handle_marshal_structure_to_ptr_with_handle() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure NativeConfig
    Public Version As Integer
    Public HandleVal As IntPtr
End Structure

Module Program
    Sub Main()
        Dim cfg As New NativeConfig With {.Version = 1, .HandleVal = New IntPtr(99)}
        Dim size = Marshal.SizeOf(GetType(NativeConfig))
        Dim mem = Marshal.AllocHGlobal(size)
        Marshal.StructureToPtr(cfg, mem, False)

        Dim readBack As NativeConfig = CType(Marshal.PtrToStructure(mem, GetType(NativeConfig)), NativeConfig)
        Marshal.FreeHGlobal(mem)

        Console.WriteLine(readBack.Version & "|" & readBack.HandleVal.ToInt64())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|99"]);
}

#[test]
fn test_vb_safe_handle_subclass_constructor_initial_state() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class FreshSafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New FreshSafeHandle()
        Console.WriteLine("IsClosed: " & h.IsClosed & "|IsInvalid: " & h.IsInvalid)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["IsClosed: False|IsInvalid: True"]);
}

#[test]
fn test_vb_safe_handle_release_handle_returning_false() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class FailingReleaseSafeHandle
    Inherits SafeHandle

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(New IntPtr(50))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Console.WriteLine("Failing Release Executed")
        Return False ' Signals failed release!
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New FailingReleaseSafeHandle()
        h.Dispose()
        Console.WriteLine("Disposed")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Failing Release Executed", "Disposed"]);
}

#[test]
fn test_vb_safe_handle_array_of_handles() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class ArrayItemHandle
    Inherits SafeHandle
    Public Sub New(val As Long)
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(New IntPtr(val))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim handles As ArrayItemHandle() = {New ArrayItemHandle(1), New ArrayItemHandle(2)}
        For Each h In handles
            h.Dispose()
        Next
        Console.WriteLine("All Array Handles Disposed")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["All Array Handles Disposed"]);
}

#[test]
fn test_vb_safe_handle_releasing_during_finalization() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class GcFinalizedSafeHandle
    Inherits SafeHandle
    Public Shared FinalizedReleaseCount As Integer = 0

    Public Sub New()
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(New IntPtr(1234))
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        FinalizedReleaseCount += 1
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Sub()
            Dim h As New GcFinalizedSafeHandle()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        Console.WriteLine("Finalized Release Count: " & GcFinalizedSafeHandle.FinalizedReleaseCount)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Finalized Release Count: 1"]);
}

#[test]
fn test_vb_safe_handle_subclass_generic_wrapper() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class GenericSafeHandle(Of T)
    Inherits SafeHandle

    Public Sub New(initialPtr As IntPtr)
        MyBase.New(IntPtr.Zero, ownsHandle:=True)
        SetHandle(initialPtr)
    End Sub

    Public Overrides ReadOnly Property IsInvalid As Boolean
        Get
            Return handle = IntPtr.Zero
        End Get
    End Property

    Protected Overrides Function ReleaseHandle() As Boolean
        Console.WriteLine("Generic Release: " & GetType(T).Name)
        Return True
    End Function
End Class

Module Program
    Sub Main()
        Dim h As New GenericSafeHandle(Of String)(New IntPtr(100))
        h.Dispose()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Generic Release: String"]);
}

#[test]
fn test_vb_safe_buffer_memory_allocation_check() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        ' SafeBuffer interop check
        Dim ptr = Marshal.AllocHGlobal(128)
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine("AllocHGlobal Safe")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["AllocHGlobal Safe"]);
}
