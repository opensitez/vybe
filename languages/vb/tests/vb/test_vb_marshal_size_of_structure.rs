use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.Runtime.InteropServices.Marshal Memory Layout
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_marshal_sizeof_primitive_types() {
    let src = r#"
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Console.WriteLine(Marshal.SizeOf(GetType(Integer)) & "|" & Marshal.SizeOf(GetType(Long)) & "|" & Marshal.SizeOf(GetType(Byte)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4|8|1"]);
}

#[test]
fn test_vb_marshal_sizeof_struct() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure NativePoint
    Public X As Integer
    Public Y As Integer
End Structure

Module Program
    Sub Main()
        Console.WriteLine(Marshal.SizeOf(GetType(NativePoint)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8"]);
}

#[test]
fn test_vb_marshal_sizeof_generic_structure() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure GenericPair(Of T)
    Public First As T
    Public Second As T
End Structure

Module Program
    Sub Main()
        Console.WriteLine(Marshal.SizeOf(GetType(GenericPair(Of Double))))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["16"]);
}

#[test]
fn test_vb_marshal_alloc_hglobal_and_free() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.AllocHGlobal(100)
        Console.WriteLine(ptr <> IntPtr.Zero)
        Marshal.FreeHGlobal(ptr)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_marshal_write_int32_read_int32() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.AllocHGlobal(4)
        Marshal.WriteInt32(ptr, 123456)
        Dim val = Marshal.ReadInt32(ptr)
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["123456"]);
}

#[test]
fn test_vb_marshal_write_byte_read_byte_sequence() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.AllocHGlobal(3)
        Marshal.WriteByte(ptr, 0, 10)
        Marshal.WriteByte(ptr, 1, 20)
        Marshal.WriteByte(ptr, 2, 30)
        Console.WriteLine(Marshal.ReadByte(ptr, 0) & "|" & Marshal.ReadByte(ptr, 1) & "|" & Marshal.ReadByte(ptr, 2))
        Marshal.FreeHGlobal(ptr)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10|20|30"]);
}

#[test]
fn test_vb_marshal_copy_byte_array_to_and_from_hglobal() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim source As Byte() = {100, 101, 102, 103}
        Dim ptr As IntPtr = Marshal.AllocHGlobal(source.Length)
        Marshal.Copy(source, 0, ptr, source.Length)

        Dim dest(3) As Byte
        Marshal.Copy(ptr, dest, 0, dest.Length)
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine(String.Join(",", dest))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["100,101,102,103"]);
}

#[test]
fn test_vb_marshal_copy_int_array_to_and_from_hglobal() {
    let src = r#"
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
        Console.WriteLine(String.Join("-", dest))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10-20-30"]);
}

#[test]
fn test_vb_marshal_structure_to_ptr_and_ptr_to_structure() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure Header
    Public Version As Integer
    Public Flag As Byte
End Structure

Module Program
    Sub Main()
        Dim h As New Header With {.Version = 2, .Flag = 1}
        Dim size = Marshal.SizeOf(GetType(Header))
        Dim ptr As IntPtr = Marshal.AllocHGlobal(size)

        Marshal.StructureToPtr(h, ptr, False)
        Dim restored As Header = CType(Marshal.PtrToStructure(ptr, GetType(Header)), Header)
        Marshal.FreeHGlobal(ptr)

        Console.WriteLine(restored.Version & "|" & restored.Flag)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|1"]);
}

#[test]
fn test_vb_marshal_string_to_hglobal_ansi_and_free() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.StringToHGlobalAnsi("NativeAnsi")
        Dim restored = Marshal.PtrToStringAnsi(ptr)
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine(restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["NativeAnsi"]);
}

#[test]
fn test_vb_marshal_string_to_hglobal_uni_and_free() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.StringToHGlobalUni("UnicodeNative")
        Dim restored = Marshal.PtrToStringUni(ptr)
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine(restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["UnicodeNative"]);
}

#[test]
fn test_vb_marshal_string_to_co_task_mem_ansi() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.StringToCoTaskMemAnsi("CoTaskAnsi")
        Dim restored = Marshal.PtrToStringAnsi(ptr)
        Marshal.FreeCoTaskMem(ptr)
        Console.WriteLine(restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["CoTaskAnsi"]);
}

#[test]
fn test_vb_marshal_string_to_bstr_and_free() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.StringToBSTR("BStrContent")
        Dim restored = Marshal.PtrToStringBSTR(ptr)
        Marshal.FreeBSTR(ptr)
        Console.WriteLine(restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["BStrContent"]);
}

#[test]
fn test_vb_marshal_get_last_win32_error_simulation() {
    let src = r#"
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim err = Marshal.GetLastWin32Error()
        Console.WriteLine(err >= 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_marshal_offset_of_struct_field() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure CompoundStruct
    Public A As Byte
    Public B As Integer
End Structure

Module Program
    Sub Main()
        ' Alignment padding: Offset of B should be 4 bytes due to 4-byte int alignment!
        Dim offsetB = Marshal.OffsetOf(GetType(CompoundStruct), "B")
        Console.WriteLine(offsetB.ToInt32())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4"]);
}

#[test]
fn test_vb_marshal_write_int64_read_int64() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.AllocHGlobal(8)
        Marshal.WriteInt64(ptr, 987654321012345L)
        Dim val = Marshal.ReadInt64(ptr)
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["987654321012345"]);
}

#[test]
fn test_vb_marshal_write_intptr_read_intptr() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim targetPtr As New IntPtr(1024)
        Dim buffer As IntPtr = Marshal.AllocHGlobal(IntPtr.Size)
        Marshal.WriteIntPtr(buffer, targetPtr)
        Dim readPtr = Marshal.ReadIntPtr(buffer)
        Marshal.FreeHGlobal(buffer)
        Console.WriteLine(readPtr.ToInt64())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1024"]);
}

#[test]
fn test_vb_marshal_get_function_pointer_for_delegate() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Delegate Function BinaryOp(a As Integer, b As Integer) As Integer

Module Program
    Private Function Add(a As Integer, b As Integer) As Integer
        Return a + b
    End Function

    Sub Main()
        Dim del As BinaryOp = AddressOf Add
        Dim funcPtr As IntPtr = Marshal.GetFunctionPointerForDelegate(del)
        Console.WriteLine(funcPtr <> IntPtr.Zero)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_marshal_get_delegate_for_function_pointer() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Delegate Function ComputeFunc(x As Integer) As Integer

Module Program
    Private Function Square(x As Integer) As Integer
        Return x * x
    End Function

    Sub Main()
        Dim delOrig As ComputeFunc = AddressOf Square
        Dim ptr = Marshal.GetFunctionPointerForDelegate(delOrig)
        Dim delRestored As ComputeFunc = CType(Marshal.GetDelegateForFunctionPointer(ptr, GetType(ComputeFunc)), ComputeFunc)
        Console.WriteLine(delRestored(5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["25"]);
}

#[test]
fn test_vb_marshal_alloc_co_task_mem_free() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.AllocCoTaskMem(64)
        Console.WriteLine(ptr <> IntPtr.Zero)
        Marshal.FreeCoTaskMem(ptr)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}
