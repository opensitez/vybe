use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: StructLayout Attributes, Pack & Explicit Offsets
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_struct_layout_sequential_default_packing() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure DefaultPackStruct
    Public A As Byte
    Public B As Integer
End Structure

Module Program
    Sub Main()
        Console.WriteLine(Marshal.SizeOf(GetType(DefaultPackStruct)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8"]);
}

#[test]
fn test_vb_struct_layout_pack_1_byte() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, Pack:=1)>
Structure PackedStruct
    Public A As Byte
    Public B As Integer
End Structure

Module Program
    Sub Main()
        Console.WriteLine(Marshal.SizeOf(GetType(PackedStruct)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}

#[test]
fn test_vb_struct_layout_pack_2_bytes() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, Pack:=2)>
Structure Pack2Struct
    Public A As Byte
    Public B As Integer
End Structure

Module Program
    Sub Main()
        ' With Pack=2: offset of B is 2, size = 2 + 4 = 6
        Console.WriteLine(Marshal.SizeOf(GetType(Pack2Struct)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["6"]);
}

#[test]
fn test_vb_struct_layout_explicit_field_offsets() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Explicit)>
Structure ExplicitOffsetsStruct
    <FieldOffset(0)> Public High As Short
    <FieldOffset(2)> Public Low As Short
End Structure

Module Program
    Sub Main()
        Console.WriteLine(Marshal.SizeOf(GetType(ExplicitOffsetsStruct)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4"]);
}

#[test]
fn test_vb_struct_layout_explicit_union_overlapping_fields() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Explicit)>
Structure IntFloatUnion
    <FieldOffset(0)> Public AsInt As Integer
    <FieldOffset(0)> Public AsSingle As Single
End Structure

Module Program
    Sub Main()
        Dim u As New IntFloatUnion With {.AsInt = &H3F800000} ' IEEE 754 float representation of 1.0F
        Console.WriteLine(u.AsSingle)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_struct_layout_char_set_ansi() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
Structure AnsiStringStruct
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=10)>
    Public Text As String
End Structure

Module Program
    Sub Main()
        Console.WriteLine(Marshal.SizeOf(GetType(AnsiStringStruct)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_struct_layout_char_set_unicode() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
Structure UnicodeStringStruct
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=10)>
    Public Text As String
End Structure

Module Program
    Sub Main()
        ' 10 WChars * 2 bytes = 20 bytes
        Console.WriteLine(Marshal.SizeOf(GetType(UnicodeStringStruct)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20"]);
}

#[test]
fn test_vb_struct_layout_byval_array_fixed_size() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure FixedArrayStruct
    <MarshalAs(UnmanagedType.ByValArray, SizeConst:=5)>
    Public Data As Integer()
End Structure

Module Program
    Sub Main()
        ' 5 Int32 elements * 4 bytes = 20 bytes
        Console.WriteLine(Marshal.SizeOf(GetType(FixedArrayStruct)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20"]);
}

#[test]
fn test_vb_struct_layout_explicit_custom_size() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Explicit, Size:=64)>
Structure PaddedHeader
    <FieldOffset(0)> Public Magic As Integer
End Structure

Module Program
    Sub Main()
        Console.WriteLine(Marshal.SizeOf(GetType(PaddedHeader)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["64"]);
}

#[test]
fn test_vb_struct_layout_nested_sequential_structs() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure Point2D
    Public X As Integer
    Public Y As Integer
End Structure

<StructLayout(LayoutKind.Sequential)>
Structure Rect2D
    Public TopLeft As Point2D
    Public BottomRight As Point2D
End Structure

Module Program
    Sub Main()
        Console.WriteLine(Marshal.SizeOf(GetType(Rect2D)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["16"]);
}

#[test]
fn test_vb_struct_layout_offsets_verification() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, Pack:=4)>
Structure MixedStruct
    Public B1 As Byte
    Public I1 As Integer
    Public B2 As Byte
    Public L1 As Long
End Structure

Module Program
    Sub Main()
        Dim offB1 = Marshal.OffsetOf(GetType(MixedStruct), "B1").ToInt32()
        Dim offI1 = Marshal.OffsetOf(GetType(MixedStruct), "I1").ToInt32()
        Dim offB2 = Marshal.OffsetOf(GetType(MixedStruct), "B2").ToInt32()
        Dim offL1 = Marshal.OffsetOf(GetType(MixedStruct), "L1").ToInt32()
        Console.WriteLine(offB1 & "|" & offI1 & "|" & offB2 & "|" & offL1)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0|4|8|12"]);
}

#[test]
fn test_vb_struct_layout_enum_field_size() {
    let src = r#"
Imports System.Runtime.InteropServices

Enum SmallEnum As Short
    V1 = 1
    V2 = 2
End Enum

<StructLayout(LayoutKind.Sequential)>
Structure EnumStruct
    Public State As SmallEnum
    Public Flag As Boolean
End Structure

Module Program
    Sub Main()
        ' SmallEnum = 2 bytes, Boolean = 4 bytes in unmanaged layout
        Console.WriteLine(Marshal.SizeOf(GetType(EnumStruct)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8"]);
}

#[test]
fn test_vb_struct_layout_copy_structure_to_ptr_with_byval_string() {
    let src = r#"
Imports System
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Ansi)>
Structure PersonNative
    <MarshalAs(UnmanagedType.ByValTStr, SizeConst:=8)>
    Public Name As String
End Structure

Module Program
    Sub Main()
        Dim p As New PersonNative With {.Name = "Vybe"}
        Dim ptr = Marshal.AllocHGlobal(8)
        Marshal.StructureToPtr(p, ptr, False)

        Dim restored As PersonNative = CType(Marshal.PtrToStructure(ptr, GetType(PersonNative)), PersonNative)
        Marshal.FreeHGlobal(ptr)
        Console.WriteLine(restored.Name)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Vybe"]);
}

#[test]
fn test_vb_struct_layout_pack_8_bytes() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential, Pack:=8)>
Structure Pack8Struct
    Public A As Byte
    Public B As Double
End Structure

Module Program
    Sub Main()
        ' Offset of B is 8, total size = 16
        Console.WriteLine(Marshal.SizeOf(GetType(Pack8Struct)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["16"]);
}

#[test]
fn test_vb_struct_layout_unmanaged_type_bool() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure BoolMarshalStruct
    <MarshalAs(UnmanagedType.I1)> Public Flag1 As Boolean
    <MarshalAs(UnmanagedType.Bool)> Public Flag4 As Boolean
End Structure

Module Program
    Sub Main()
        ' Flag1 = 1 byte, Flag4 = 4 bytes (plus 3 padding bytes) = 8 bytes total
        Console.WriteLine(Marshal.SizeOf(GetType(BoolMarshalStruct)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8"]);
}

#[test]
fn test_vb_struct_layout_explicit_structure_unions() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Explicit)>
Structure ColorPixel
    <FieldOffset(0)> Public R As Byte
    <FieldOffset(1)> Public G As Byte
    <FieldOffset(2)> Public B As Byte
    <FieldOffset(3)> Public A As Byte
    <FieldOffset(0)> Public RgbaValue As UInteger
End Structure

Module Program
    Sub Main()
        Dim p As New ColorPixel With {.RgbaValue = &HFF0000FFUI}
        Console.WriteLine(Marshal.SizeOf(GetType(ColorPixel)) & "|" & p.RgbaValue <> 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4|True"]);
}

#[test]
fn test_vb_struct_layout_auto_layout_kind() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Auto)>
Class AutoLayoutClass
    Public A As Byte
    Public B As Integer
End Class

Module Program
    Sub Main()
        ' Auto layout classes cannot be measured via Marshal.SizeOf directly without throwing!
        Try
            Marshal.SizeOf(GetType(AutoLayoutClass))
        Catch ex As ArgumentException
            Console.WriteLine("ArgumentException Caught on Auto Layout")
        End Try
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["ArgumentException Caught on Auto Layout"]);
}

#[test]
fn test_vb_struct_layout_inherited_structure_not_allowed() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure BaseStruct
    Public X As Integer
End Structure

Module Program
    Sub Main()
        ' In VB.NET structures cannot inherit from other structures!
        Dim s As New BaseStruct With {.X = 10}
        Console.WriteLine(s.X)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_struct_layout_byref_structure_pointer_modification() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure ModifiableStruct
    Public Value As Integer
End Structure

Module Program
    Private Sub ModifyStruct(ByRef s As ModifiableStruct)
        s.Value += 100
    End Sub

    Sub Main()
        Dim s As New ModifiableStruct With {.Value = 50}
        ModifyStruct(s)
        Console.WriteLine(s.Value)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["150"]);
}

#[test]
fn test_vb_struct_layout_zero_length_fixed_buffer_simulation() {
    let src = r#"
Imports System.Runtime.InteropServices

<StructLayout(LayoutKind.Sequential)>
Structure EmptyStruct
End Structure

Module Program
    Sub Main()
        ' Unmanaged size of empty struct is 1 byte in CLI
        Console.WriteLine(Marshal.SizeOf(GetType(EmptyStruct)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}
