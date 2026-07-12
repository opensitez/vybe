//! `BitConverter` maps primitive values to bytes and back on the host endianness.
use super::helpers::run_csharp;

#[test]
fn bit_converter_int32_roundtrip_preserves_numeric_value() {
    assert_eq!(
        run_csharp(
            r#"
var bytes = System.BitConverter.GetBytes(1024);
Console.WriteLine(System.BitConverter.ToInt32(bytes, 0));
"#
        ),
        &["1024"]
    );
}

#[test]
fn bit_converter_double_bytes_reconstruct_original_fraction() {
    assert_eq!(
        run_csharp(
            r#"
var bytes = System.BitConverter.GetBytes(2.5);
Console.WriteLine(System.BitConverter.ToDouble(bytes, 0));
"#
        ),
        &["2.5"]
    );
}

#[test]
fn bit_converter_is_little_endian_flag_matches_platform_expectation() {
    assert_eq!(
        run_csharp(
            r#"
Console.WriteLine(System.BitConverter.IsLittleEndian);
"#
        ),
        &["True"]
    );
}

#[test]
fn bit_array_set_and_get_roundtrip_single_index() {
    assert_eq!(
        run_csharp(
            r#"
var bits = new System.Collections.BitArray(3);
bits[1] = true;
Console.WriteLine(bits[1]);
Console.WriteLine(bits[0]);
"#
        ),
        &["True", "False"]
    );
}

#[test]
fn bit_array_and_applies_elementwise_conjunction() {
    assert_eq!(
        run_csharp(
            r#"
var left = new System.Collections.BitArray(new bool[] { true, false, true });
var right = new System.Collections.BitArray(new bool[] { true, true, false });
left.And(right);
Console.WriteLine(left[0]);
Console.WriteLine(left[1]);
"#
        ),
        &["True", "False"]
    );
}

#[test]
fn bit_array_not_inverts_all_bits() {
    assert_eq!(
        run_csharp(
            r#"
var bits = new System.Collections.BitArray(new bool[] { false, true });
bits.Not();
Console.WriteLine(bits[0]);
Console.WriteLine(bits[1]);
"#
        ),
        &["True", "False"]
    );
}
