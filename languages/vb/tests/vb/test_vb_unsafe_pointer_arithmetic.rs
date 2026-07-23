use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Memory Pointer Offsets & Safe Managed Buffer Pointers
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_pointer_offset_calculation_byte_array() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim bytes As Byte() = {10, 20, 30, 40, 50}
        Dim handle = GCHandle.Alloc(bytes, GCHandleType.Pinned)
        Dim baseAddr = handle.AddrOfPinnedObject()

        Dim thirdByteAddr = IntPtr.Add(baseAddr, 2)
        Dim val = Marshal.ReadByte(thirdByteAddr)
        handle.Free()
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30"]);
}

#[test]
fn test_vb_pointer_offset_calculation_int32_array() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim numbers As Integer() = {100, 200, 300, 400}
        Dim handle = GCHandle.Alloc(numbers, GCHandleType.Pinned)
        Dim baseAddr = handle.AddrOfPinnedObject()

        ' Offset for element index 2: 2 * Marshal.SizeOf(GetType(Integer)) = 8 bytes
        Dim elem2Addr = IntPtr.Add(baseAddr, 2 * 4)
        Dim val = Marshal.ReadInt32(elem2Addr)
        handle.Free()
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["300"]);
}

#[test]
fn test_vb_pointer_offset_write_int32() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim numbers As Integer() = {0, 0, 0}
        Dim handle = GCHandle.Alloc(numbers, GCHandleType.Pinned)
        Dim baseAddr = handle.AddrOfPinnedObject()

        Marshal.WriteInt32(IntPtr.Add(baseAddr, 0), 10)
        Marshal.WriteInt32(IntPtr.Add(baseAddr, 4), 20)
        Marshal.WriteInt32(IntPtr.Add(baseAddr, 8), 30)
        handle.Free()
        Console.WriteLine(String.Join(",", numbers))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10,20,30"]);
}

#[test]
fn test_vb_pointer_difference_calculation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim p1 As New IntPtr(2000)
        Dim p2 As New IntPtr(1000)
        Dim diff As Long = p1.ToInt64() - p2.ToInt64()
        Console.WriteLine(diff)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1000"]);
}

#[test]
fn test_vb_pointer_structure_array_indexing() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure Element
    Public ID As Integer
    Public Value As Double
End Structure

Module Program
    Sub Main()
        Dim elements(1) As Element
        elements(0) = New Element With {.ID = 1, .Value = 1.1}
        elements(1) = New Element With {.ID = 2, .Value = 2.2}

        Dim handle = GCHandle.Alloc(elements, GCHandleType.Pinned)
        Dim baseAddr = handle.AddrOfPinnedObject()

        Dim elemSize = Marshal.SizeOf(GetType(Element))
        Dim secondElemAddr = IntPtr.Add(baseAddr, elemSize)
        Dim restored2 As Element = CType(Marshal.PtrToStructure(secondElemAddr, GetType(Element)), Element)
        handle.Free()

        Console.WriteLine(restored2.ID & ":" & restored2.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2:2.2"]);
}

#[test]
fn test_vb_pointer_unmanaged_memory_buffer_iteration() {
    let src = r#"
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
"#;
    assert_eq!(run_vb(src), vec!["10,20,30,40,50"]);
}

#[test]
fn test_vb_pointer_aligned_offset_checks() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim p As New IntPtr(1024)
        Dim isAligned4 = (p.ToInt64() Mod 4L = 0)
        Dim isAligned8 = (p.ToInt64() Mod 8L = 0)
        Console.WriteLine(isAligned4 & "|" & isAligned8)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|True"]);
}

#[test]
fn test_vb_pointer_string_null_terminated_reader() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.StringToHGlobalAnsi("NullTerminated")
        Dim str = Marshal.PtrToStringAnsi(ptr)
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine(str)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["NullTerminated"]);
}

#[test]
fn test_vb_pointer_string_length_bounded_reader() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.StringToHGlobalAnsi("SubstringText")
        ' Read first 9 characters only
        Dim str = Marshal.PtrToStringAnsi(ptr, 9)
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine(str)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Substring"]);
}

#[test]
fn test_vb_pointer_hglobal_realloc_expansion() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.AllocHGlobal(4)
        Marshal.WriteInt32(ptr, 99)

        Dim newPtr As IntPtr = Marshal.ReAllocHGlobal(ptr, CType(8, IntPtr))
        Dim val1 = Marshal.ReadInt32(newPtr)
        Marshal.WriteInt32(IntPtr.Add(newPtr, 4), 100)
        Dim val2 = Marshal.ReadInt32(IntPtr.Add(newPtr, 4))
        Marshal.FreeHGlobal(newPtr)

        Console.WriteLine(val1 & "|" & val2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99|100"]);
}

#[test]
fn test_vb_pointer_copy_memory_between_pointers() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim srcPtr As IntPtr = Marshal.AllocHGlobal(4)
        Dim destPtr As IntPtr = Marshal.AllocHGlobal(4)

        Marshal.WriteInt32(srcPtr, 777)
        Dim tempArr(3) As Byte
        Marshal.Copy(srcPtr, tempArr, 0, 4)
        Marshal.Copy(tempArr, 0, destPtr, 4)

        Dim copiedVal = Marshal.ReadInt32(destPtr)
        Marshal.FreeHGlobal(srcPtr)
        Marshal.FreeHGlobal(destPtr)
        Console.WriteLine(copiedVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["777"]);
}

#[test]
fn test_vb_pointer_zero_fill_memory_buffer() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.AllocHGlobal(4)
        Marshal.WriteInt32(ptr, &HFFFFFFFF)

        ' Overwrite with zeros
        Dim zeros(3) As Byte
        Marshal.Copy(zeros, 0, ptr, 4)
        Dim val = Marshal.ReadInt32(ptr)
        Marshal.FreeHGlobal(ptr)

        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_pointer_unmanaged_type_size_computations() {
    let src = r#"
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim charSize = Marshal.SystemDefaultCharSize
        Console.WriteLine(charSize = 1 OrElse charSize = 2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_pointer_span_from_intptr_read_bytes() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim data As Byte() = {5, 10, 15, 20}
        Dim handle = GCHandle.Alloc(data, GCHandleType.Pinned)
        Dim ptr = handle.AddrOfPinnedObject()

        Dim span As ReadOnlySpan(Of Byte) = New ReadOnlySpan(Of Byte)(ptr.ToPointer(), 4)
        Console.WriteLine(span(0) & "|" & span(3))
        handle.Free()
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5|20"]);
}

#[test]
fn test_vb_pointer_read_int16_short_values() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.AllocHGlobal(4)
        Marshal.WriteInt16(ptr, 0, CShort(1000))
        Marshal.WriteInt16(ptr, 2, CShort(2000))

        Dim s1 = Marshal.ReadInt16(ptr, 0)
        Dim s2 = Marshal.ReadInt16(ptr, 2)
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine(s1 & "|" & s2)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1000|2000"]);
}

#[test]
fn test_vb_pointer_read_int64_long_values() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.AllocHGlobal(8)
        Marshal.WriteInt64(ptr, 0, 5000000000L)
        Dim l1 = Marshal.ReadInt64(ptr, 0)
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine(l1)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5000000000"]);
}

#[test]
fn test_vb_pointer_read_single_float_values() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.AllocHGlobal(4)
        Dim f As Single = 12.34F
        Dim bits = BitConverter.SingleToInt32Bits(f)
        Marshal.WriteInt32(ptr, bits)

        Dim readBits = Marshal.ReadInt32(ptr)
        Dim restoredF = BitConverter.Int32BitsToSingle(readBits)
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine(restoredF)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["12.34"]);
}

#[test]
fn test_vb_pointer_read_double_float_values() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.AllocHGlobal(8)
        Dim d As Double = 99.87654
        Dim bits = BitConverter.DoubleToInt64Bits(d)
        Marshal.WriteInt64(ptr, bits)

        Dim readBits = Marshal.ReadInt64(ptr)
        Dim restoredD = BitConverter.Int64BitsToDouble(readBits)
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine(restoredD)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["99.87654"]);
}

#[test]
fn test_vb_pointer_secure_string_conversion_simulation() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices
Imports System.Security

Module Program
    Sub Main()
        Dim ss As New SecureString()
        ss.AppendChar("A"c)
        ss.AppendChar("B"c)
        Dim ptr As IntPtr = Marshal.SecureStringToGlobalAllocUnicode(ss)
        Dim str = Marshal.PtrToStringUni(ptr)
        Marshal.ZeroFreeGlobalAllocUnicode(ptr)
        ss.Dispose()
        Console.WriteLine(str)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["AB"]);
}

#[test]
fn test_vb_pointer_unmanaged_memory_stream_wrapper() {
    let src = r#"
Imports System
Imports System.IO
Imports System.Runtime.InteropServices

Module Program
    Sub Main()
        Dim ptr As IntPtr = Marshal.AllocHGlobal(100)
        Using ums As New UnmanagedMemoryStream(CType(ptr.ToPointer(), Byte*), 100, 100, FileAccess.ReadWrite)
            ums.WriteByte(77)
            ums.Position = 0
            Console.WriteLine(ums.ReadByte())
        End Using
        Marshal.FreeHGlobal(ptr)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["77"]);
}
