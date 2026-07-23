use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: System.BitConverter Primitive Binary Conversion
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_bitconverter_get_bytes_int32() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes = BitConverter.GetBytes(1024)
        Dim restored = BitConverter.ToInt32(bytes, 0)
        Console.WriteLine(bytes.Length & "|" & restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4|1024"]);
}

#[test]
fn test_vb_bitconverter_get_bytes_double() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim orig As Double = 3.141592653589793
        Dim bytes = BitConverter.GetBytes(orig)
        Dim restored = BitConverter.ToDouble(bytes, 0)
        Console.WriteLine(bytes.Length & "|" & (orig = restored))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["8|True"]);
}

#[test]
fn test_vb_bitconverter_get_bytes_boolean() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytesTrue = BitConverter.GetBytes(True)
        Dim bytesFalse = BitConverter.GetBytes(False)
        Console.WriteLine(bytesTrue(0) & "|" & bytesFalse(0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1|0"]);
}

#[test]
fn test_vb_bitconverter_get_bytes_char_unicode() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes = BitConverter.GetBytes("Z"c)
        Dim restored = BitConverter.ToChar(bytes, 0)
        Console.WriteLine(bytes.Length & "|" & restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|Z"]);
}

#[test]
fn test_vb_bitconverter_is_little_endian_check() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Console.WriteLine(BitConverter.IsLittleEndian)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_bitconverter_to_string_hyphenated_hex() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes As Byte() = {10, 20, 30, 40}
        Dim hexStr = BitConverter.ToString(bytes)
        Console.WriteLine(hexStr)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0A-14-1E-28"]);
}

#[test]
fn test_vb_bitconverter_to_string_subslice() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes As Byte() = {255, 128, 64, 32, 16}
        ' Subslice starting at index 1 for length 3
        Dim hexStr = BitConverter.ToString(bytes, 1, 3)
        Console.WriteLine(hexStr)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["80-40-20"]);
}

#[test]
fn test_vb_bitconverter_double_to_int64_bits_roundtrip() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim dbl = 2.718281828459
        Dim bits = BitConverter.DoubleToInt64Bits(dbl)
        Dim restored = BitConverter.Int64BitsToDouble(bits)
        Console.WriteLine(dbl = restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_bitconverter_single_to_int32_bits_roundtrip() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim flt As Single = 123.45F
        Dim bits = BitConverter.SingleToInt32Bits(flt)
        Dim restored = BitConverter.Int32BitsToSingle(bits)
        Console.WriteLine(flt = restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_bitconverter_to_int16_short() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes As Byte() = BitConverter.GetBytes(CShort(-30000))
        Dim restored = BitConverter.ToInt16(bytes, 0)
        Console.WriteLine(restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-30000"]);
}

#[test]
fn test_vb_bitconverter_to_uint16_ushort() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes As Byte() = BitConverter.GetBytes(65000US)
        Dim restored = BitConverter.ToUInt16(bytes, 0)
        Console.WriteLine(restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["65000"]);
}

#[test]
fn test_vb_bitconverter_to_uint32_uinteger() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes As Byte() = BitConverter.GetBytes(4000000000UI)
        Dim restored = BitConverter.ToUInt32(bytes, 0)
        Console.WriteLine(restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["4000000000"]);
}

#[test]
fn test_vb_bitconverter_to_uint64_ulong() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes As Byte() = BitConverter.GetBytes(18000000000000000000UL)
        Dim restored = BitConverter.ToUInt64(bytes, 0)
        Console.WriteLine(restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["18000000000000000000"]);
}

#[test]
fn test_vb_bitconverter_to_int64_long() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes As Byte() = BitConverter.GetBytes(-9000000000000000000L)
        Dim restored = BitConverter.ToInt64(bytes, 0)
        Console.WriteLine(restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["-9000000000000000000"]);
}

#[test]
fn test_vb_bitconverter_to_boolean() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bTrue = BitConverter.ToBoolean(New Byte() {1}, 0)
        Dim bFalse = BitConverter.ToBoolean(New Byte() {0}, 0)
        Console.WriteLine(bTrue & "|" & bFalse)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_bitconverter_out_of_range_start_index_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim bytes As Byte() = {1, 2, 3}
        Try
            BitConverter.ToInt32(bytes, 1) ' Needs 4 bytes, only 2 left!
        Catch ex As ArgumentException
            Console.WriteLine("ArgumentException Caught on Truncated Buffer")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentException Caught on Truncated Buffer"]
    );
}

#[test]
fn test_vb_bitconverter_null_buffer_throws() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Try
            BitConverter.ToInt32(Nothing, 0)
        Catch ex As ArgumentNullException
            Console.WriteLine("ArgumentNullException Caught on Null Buffer")
        End Try
    End Sub
End Module
"#;
    assert_eq!(
        run_vb(src),
        vec!["ArgumentNullException Caught on Null Buffer"]
    );
}

#[test]
fn test_vb_bitconverter_half_float_conversion() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim h As Half = CType(1.5F, Half)
        Dim bytes = BitConverter.GetBytes(h)
        Dim restored = BitConverter.ToHalf(bytes, 0)
        Console.WriteLine(bytes.Length & "|" & restored)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2|1.5"]);
}

#[test]
fn test_vb_bitconverter_span_overloads() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim span As ReadOnlySpan(Of Byte) = New Byte() {10, 0, 0, 0}
        Dim val = BitConverter.ToInt32(span)
        Console.WriteLine(val)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_bitconverter_try_write_bytes_span() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim destination(3) As Byte
        Dim span As Span(Of Byte) = destination
        Dim ok = BitConverter.TryWriteBytes(span, 9999)
        Console.WriteLine(ok & "|" & BitConverter.ToInt32(destination, 0))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|9999"]);
}
