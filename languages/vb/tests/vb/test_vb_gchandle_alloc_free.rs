use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Runtime.InteropServices.GCHandle Management
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_gchandle_alloc_normal_and_free() {
    let src = r#"
Imports System.Runtime.InteropServices

Class TargetObj
    Public Value As Integer = 42
End Class

Module Program
    Sub Main()
        Dim obj As New TargetObj()
        Dim handle As GCHandle = GCHandle.Alloc(obj, GCHandleType.Normal)
        Dim retrieved As TargetObj = CType(handle.Target, TargetObj)
        Console.WriteLine(handle.IsAllocated & "|" & retrieved.Value)
        handle.Free()
        Console.WriteLine(handle.IsAllocated)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|42", "False"]);
}

#[test]
fn test_vb_gchandle_alloc_pinned_byte_array() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim bytes As Byte() = {10, 20, 30, 40}
        Dim handle As GCHandle = GCHandle.Alloc(bytes, GCHandleType.Pinned)
        Dim addr As IntPtr = handle.AddrOfPinnedObject()
        Console.WriteLine((addr <> IntPtr.Zero) & "|" & Marshal.ReadByte(addr))
        handle.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|10"]);
}

#[test]
fn test_vb_gchandle_alloc_weak_track_resurrection() {
    let src = r#"
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim obj As New Object()
        Dim handle As GCHandle = GCHandle.Alloc(obj, GCHandleType.WeakTrackResurrection)
        Console.WriteLine(handle.IsAllocated & "|" & (handle.Target IsNot Nothing))
        handle.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_gchandle_to_intptr_and_from_intptr_roundtrip() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class Container
    Public Title As String = "VybeHandle"
End Class

Module Program
    Sub Main()
        Dim c As New Container()
        Dim handle = GCHandle.Alloc(c)
        Dim ptr As IntPtr = GCHandle.ToIntPtr(handle)

        Dim restoredHandle = GCHandle.FromIntPtr(ptr)
        Dim restored As Container = CType(restoredHandle.Target, Container)
        Console.WriteLine(restored.Title)
        handle.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["VybeHandle"]);
}

#[test]
fn test_vb_gchandle_pinned_int32_array_address_pointer() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim numbers As Integer() = {100, 200, 300}
        Dim handle = GCHandle.Alloc(numbers, GCHandleType.Pinned)
        Dim addr = handle.AddrOfPinnedObject()
        Dim secondVal = Marshal.ReadInt32(addr, 4) ' 4 byte offset for second int
        handle.Free()
        Console.WriteLine(secondVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["200"]);
}

#[test]
fn test_vb_gchandle_op_explicit_intptr_conversions() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim target As New Object()
        Dim h = GCHandle.Alloc(target)
        Dim ptr As IntPtr = CType(h, IntPtr)
        Dim h2 As GCHandle = CType(ptr, GCHandle)
        Console.WriteLine(Object.ReferenceEquals(target, h2.Target))
        h.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gchandle_equality_operator() {
    let src = r#"
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim obj As New Object()
        Dim h1 = GCHandle.Alloc(obj)
        Dim h2 = h1
        Console.WriteLine(h1 = h2)
        h1.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gchandle_inequality_operator() {
    let src = r#"
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim obj1 As New Object()
        Dim obj2 As New Object()
        Dim h1 = GCHandle.Alloc(obj1)
        Dim h2 = GCHandle.Alloc(obj2)
        Console.WriteLine(h1 <> h2)
        h1.Free()
        h2.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gchandle_pinned_non_blittable_type_throws() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Class NonBlittableClass
    Public Name As String
End Class

Module Program
    Sub Main()
        Dim obj As New NonBlittableClass With {.Name = "Test"}
        Try
            Dim handle = GCHandle.Alloc(obj, GCHandleType.Pinned)
        Catch ex As ArgumentException
            Console.WriteLine("ArgumentException Caught on Pinning Non-Blittable")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentException Caught on Pinning Non-Blittable"]
    );
}

#[test]
fn test_vb_gchandle_addr_of_pinned_object_on_unpinned_throws() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim obj As New Object()
        Dim handle = GCHandle.Alloc(obj, GCHandleType.Normal)
        Try
            handle.AddrOfPinnedObject()
        Catch ex As InvalidOperationException
            Console.WriteLine("InvalidOperationException Caught on Unpinned AddrOfPinnedObject")
        Finally
            handle.Free()
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidOperationException Caught on Unpinned AddrOfPinnedObject"]
    );
}

#[test]
fn test_vb_gchandle_free_unallocated_handle_throws() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim handle As GCHandle
        Try
            handle.Free()
        Catch ex As InvalidOperationException
            Console.WriteLine("InvalidOperationException Caught on Free Unallocated Handle")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidOperationException Caught on Free Unallocated Handle"]
    );
}

#[test]
fn test_vb_gchandle_double_free_throws() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim obj As New Object()
        Dim handle = GCHandle.Alloc(obj)
        handle.Free()
        Try
            handle.Free()
        Catch ex As InvalidOperationException
            Console.WriteLine("InvalidOperationException Caught on Double Free")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["InvalidOperationException Caught on Double Free"]
    );
}

#[test]
fn test_vb_gchandle_target_property_mutation() {
    let src = r#"
Imports System.Runtime.InteropServices

Class Holder
    Public Data As String = "Original"
End Class

Module Program
    Sub Main()
        Dim h1 As New Holder()
        Dim h2 As New Holder With {.Data = "Replaced"}
        Dim handle = GCHandle.Alloc(h1)
        handle.Target = h2

        Dim current As Holder = CType(handle.Target, Holder)
        Console.WriteLine(current.Data)
        handle.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Replaced"]);
}

#[test]
fn test_vb_gchandle_hash_code_consistency() {
    let src = r#"
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim obj As New Object()
        Dim handle = GCHandle.Alloc(obj)
        Dim hash1 = handle.GetHashCode()
        Dim hash2 = handle.GetHashCode()
        Console.WriteLine(hash1 = hash2)
        handle.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gchandle_pinned_blittable_struct() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure Vector3
    Public X As Single
    Public Y As Single
    Public Z As Single
End Structure

Module Program
    Sub Main()
        Dim v As New Vector3 With {.X = 1.0F, .Y = 2.0F, .Z = 3.0F}
        Dim handle = GCHandle.Alloc(v, GCHandleType.Pinned)
        Dim addr = handle.AddrOfPinnedObject()
        Console.WriteLine(addr <> IntPtr.Zero)
        handle.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gchandle_weak_handle_target_access() {
    let src = r#"
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim obj As New Object()
        Dim handle = GCHandle.Alloc(obj, GCHandleType.Weak)
        Console.WriteLine(handle.Target IsNot Nothing)
        handle.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_gchandle_multiple_handles_same_object() {
    let src = r#"
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim obj As New Object()
        Dim h1 = GCHandle.Alloc(obj)
        Dim h2 = GCHandle.Alloc(obj)
        Console.WriteLine((h1 <> h2) & "|" & Object.ReferenceEquals(h1.Target, h2.Target))
        h1.Free()
        h2.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_gchandle_pinned_string_unicode_chars() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim str = "PinnedString"
        Dim handle = GCHandle.Alloc(str, GCHandleType.Pinned)
        Dim addr = handle.AddrOfPinnedObject()
        Dim firstChar As Char = ChrW(Marshal.ReadInt16(addr))
        handle.Free()
        Console.WriteLine(firstChar)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["P"]);
}

#[test]
fn test_vb_gchandle_null_target_allocation() {
    let src = r#"
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim handle = GCHandle.Alloc(Nothing)
        Console.WriteLine(handle.IsAllocated & "|" & (handle.Target Is Nothing))
        handle.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_gchandle_struct_value_boxed_on_normal_alloc() {
    let src = r#"
Imports System.Runtime.InteropServices

Structure ValueHolder
    Public Count As Integer
End Structure

Module Program
    Sub Main()
        Dim v As New ValueHolder With {.Count = 99}
        Dim handle = GCHandle.Alloc(v) ' Boxing occurs for value types!
        Dim boxed As ValueHolder = CType(handle.Target, ValueHolder)
        Console.WriteLine(boxed.Count)
        handle.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99"]);
}
