use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Bitwise Shift (<<, >>) & Unsigned Integer Arithmetic
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_bitwise_shift_left_integer() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Integer = 1
        Dim res = val << 4
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["16"]);
}

#[test]
fn test_vb_bitwise_shift_right_integer() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Integer = 64
        Dim res = val >> 3
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8"]);
}

#[test]
fn test_vb_bitwise_shift_left_uinteger() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As UInteger = 1UI
        Dim res = val << 31
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2147483648"]);
}

#[test]
fn test_vb_bitwise_shift_right_unsigned_logical_shift() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As UInteger = &HFFFFFFFFUI
        Dim res = val >> 1
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2147483647"]);
}

#[test]
fn test_vb_bitwise_shift_left_ulong() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As ULong = 1UL
        Dim res = val << 40
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1099511627776"]);
}

#[test]
fn test_vb_bitwise_and_masking() {
    let src = r#"
Module Program
    Sub Main()
        Dim flags As Integer = &HF5 ' 1111 0101
        Dim mask As Integer = &HF0  ' 1111 0000
        Dim masked = flags And mask
        Console.WriteLine(Hex(masked))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["F0"]);
}

#[test]
fn test_vb_bitwise_or_flag_combination() {
    let src = r#"
Module Program
    Sub Main()
        Dim flagA As Integer = &H01
        Dim flagB As Integer = &H04
        Dim combined = flagA Or flagB
        Console.WriteLine(combined)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["5"]);
}

#[test]
fn test_vb_bitwise_xor_flag_toggle() {
    let src = r#"
Module Program
    Sub Main()
        Dim flags As Integer = &H05 ' 0101
        Dim toggle As Integer = &H01 ' 0001
        Dim toggled = flags Xor toggle ' 0100 (4)
        Console.WriteLine(toggled)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4"]);
}

#[test]
fn test_vb_bitwise_not_complement_integer() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Integer = 0
        Dim comp = Not val
        Console.WriteLine(comp)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-1"]);
}

#[test]
fn test_vb_bitwise_not_complement_uinteger() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As UInteger = 0UI
        Dim comp = Not val
        Console.WriteLine(comp)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4294967295"]);
}

#[test]
fn test_vb_bitwise_shift_left_overflow_wrap() {
    let src = r#"
Module Program
    Sub Main()
        Dim val As Byte = &H80
        Dim res As Byte = CByte(val << 1)
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_unsigned_byte_arithmetic() {
    let src = r#"
Module Program
    Sub Main()
        Dim b1 As Byte = 200
        Dim b2 As Byte = 55
        Dim sum As Byte = CByte(b1 + b2)
        Console.WriteLine(sum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["255"]);
}

#[test]
fn test_vb_unsigned_ushort_arithmetic() {
    let src = r#"
Module Program
    Sub Main()
        Dim s1 As UShort = 60000US
        Dim s2 As UShort = 5000US
        Dim sum As UShort = CUShort(s1 + s2)
        Console.WriteLine(sum)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["65000"]);
}

#[test]
fn test_vb_bitwise_rotate_left_simulation() {
    let src = r#"
Module Program
    Private Function RotateLeft(val As UInteger, shift As Integer) As UInteger
        Return (val << shift) Or (val >> (32 - shift))
    End Function

    Sub Main()
        Dim val As UInteger = &H80000001UI
        Dim rot = RotateLeft(val, 1)
        Console.WriteLine(Hex(rot))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3"]);
}

#[test]
fn test_vb_bitwise_rotate_right_simulation() {
    let src = r#"
Module Program
    Private Function RotateRight(val As UInteger, shift As Integer) As UInteger
        Return (val >> shift) Or (val << (32 - shift))
    End Function

    Sub Main()
        Dim val As UInteger = &H00000001UI
        Dim rot = RotateRight(val, 1)
        Console.WriteLine(Hex(rot))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["80000000"]);
}

#[test]
fn test_vb_bitwise_extract_bit_field() {
    let src = r#"
Module Program
    Sub Main()
        Dim packedData As UInteger = &HAABBCCDDUI
        ' Extract second byte (CC = 204)
        Dim secondByte = (packedData >> 8) And &HFFUI
        Console.WriteLine(Hex(secondByte))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["CC"]);
}

#[test]
fn test_vb_bitwise_pack_four_bytes_into_uint() {
    let src = r#"
Module Program
    Sub Main()
        Dim b1 As Byte = &HA
        Dim b2 As Byte = &HB
        Dim b3 As Byte = &HC
        Dim b4 As Byte = &HD
        Dim packed As UInteger = (CUInt(b1) << 24) Or (CUInt(b2) << 16) Or (CUInt(b3) << 8) Or CUInt(b4)
        Console.WriteLine(Hex(packed))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A0B0C0D"]);
}

#[test]
fn test_vb_bitwise_count_set_bits_popcount() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim val As UInteger = &H0F0F0F0FUI
        Dim count = BitOperations.PopCount(val)
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["16"]);
}

#[test]
fn test_vb_bitwise_trailing_zero_count() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim val As UInteger = 16UI ' 10000 (4 zeros)
        Dim count = BitOperations.TrailingZeroCount(val)
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4"]);
}

#[test]
fn test_vb_bitwise_leading_zero_count() {
    let src = r#"
Imports System.Numerics

Module Program
    Sub Main()
        Dim val As UInteger = 1UI
        Dim count = BitOperations.LeadingZeroCount(val)
        Console.WriteLine(count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["31"]);
}
