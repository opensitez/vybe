use crate::helpers::{run_in_main, run_main};

#[test]
fn byte_literal_positive_value() {
    let out = run_main("byte b = 5; System.out.println(b);");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn byte_literal_negative_value() {
    let out = run_main("byte b = -3; System.out.println(b);");
    assert_eq!(out, vec!["-3"]);
}

#[test]
fn short_literal_positive_value() {
    let out = run_main("short s = 300; System.out.println(s);");
    assert_eq!(out, vec!["300"]);
}

#[test]
fn short_literal_negative_value() {
    let out = run_main("short s = -120; System.out.println(s);");
    assert_eq!(out, vec!["-120"]);
}

#[test]
fn byte_addition_within_range() {
    let out = run_main("byte a = 10; byte b = 15; System.out.println((byte)(a + b));");
    assert_eq!(out, vec!["25"]);
}

#[test]
fn byte_addition_overflow_wraps_to_negative() {
    let out = run_main("byte a = 127; byte b = 1; System.out.println((byte)(a + b));");
    assert_eq!(out, vec!["-128"]);
}

#[test]
fn byte_subtraction_underflow_wraps_to_positive() {
    let out = run_main("byte a = -128; byte b = 1; System.out.println((byte)(a - b));");
    assert_eq!(out, vec!["127"]);
}

#[test]
fn short_addition_overflow_wraps() {
    let out = run_main("short a = 32767; short b = 1; System.out.println((short)(a + b));");
    assert_eq!(out, vec!["-32768"]);
}

#[test]
fn short_subtraction_underflow_wraps() {
    let out = run_main("short a = -32768; short b = 1; System.out.println((short)(a - b));");
    assert_eq!(out, vec!["32767"]);
}

#[test]
fn byte_multiplication_small_factors() {
    let out = run_main("byte a = 6; byte b = 7; System.out.println((byte)(a * b));");
    assert_eq!(out, vec!["42"]);
}

#[test]
fn short_multiplication_wraps_on_overflow() {
    let out = run_main("short a = 200; short b = 200; System.out.println((short)(a * b));");
    assert_eq!(out, vec!["-7920"]);
}

#[test]
fn byte_division_truncates_toward_zero() {
    let out = run_main("byte a = 7; byte b = 2; System.out.println((byte)(a / b));");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn short_division_negative_truncation() {
    let out = run_main("short a = -7; short b = 2; System.out.println((short)(a / b));");
    assert_eq!(out, vec!["-3"]);
}

#[test]
fn byte_modulo_remainder() {
    let out = run_main("byte a = 17; byte b = 5; System.out.println((byte)(a % b));");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn short_modulo_negative_operands() {
    let out = run_main("short a = -17; short b = 5; System.out.println((short)(a % b));");
    assert_eq!(out, vec!["-2"]);
}

#[test]
fn cast_int_to_byte_truncates_high_bits() {
    let out = run_main("int n = 130; byte b = (byte) n; System.out.println(b);");
    assert_eq!(out, vec!["-126"]);
}

#[test]
fn cast_int_to_short_truncates_high_bits() {
    let out = run_main("int n = 70000; short s = (short) n; System.out.println(s);");
    assert_eq!(out, vec!["4464"]);
}

#[test]
fn cast_byte_to_int_sign_extends() {
    let out = run_main("byte b = -1; int n = b; System.out.println(n);");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn cast_short_to_int_sign_extends() {
    let out = run_main("short s = -5; int n = s; System.out.println(n);");
    assert_eq!(out, vec!["-5"]);
}

#[test]
fn cast_byte_to_short_widening_within_short_range() {
    let out = run_main("byte b = 100; short s = b; System.out.println(s);");
    assert_eq!(out, vec!["100"]);
}

#[test]
fn cast_short_to_byte_narrows_with_truncation() {
    let out = run_main("short s = 300; byte b = (byte) s; System.out.println(b);");
    assert_eq!(out, vec!["44"]);
}

#[test]
fn byte_increment_wraps_past_max() {
    let out = run_main("byte b = 127; b++; System.out.println(b);");
    assert_eq!(out, vec!["-128"]);
}

#[test]
fn byte_decrement_wraps_past_min() {
    let out = run_main("byte b = -128; b--; System.out.println(b);");
    assert_eq!(out, vec!["127"]);
}

#[test]
fn short_increment_wraps_past_max() {
    let out = run_main("short s = 32767; s++; System.out.println(s);");
    assert_eq!(out, vec!["-32768"]);
}

#[test]
fn short_decrement_wraps_past_min() {
    let out = run_main("short s = -32768; s--; System.out.println(s);");
    assert_eq!(out, vec!["32767"]);
}

#[test]
fn byte_bitwise_and_masks_bits() {
    let out =
        run_main("byte a = 0b11110000; byte b = 0b10101010; System.out.println((byte)(a & b));");
    assert_eq!(out, vec!["-128"]);
}

#[test]
fn short_bitwise_or_combines_bits() {
    let out = run_main("short a = 1; short b = 2; System.out.println((short)(a | b));");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn byte_bitwise_xor_toggles_bits() {
    let out = run_main("byte a = 5; byte b = 3; System.out.println((byte)(a ^ b));");
    assert_eq!(out, vec!["6"]);
}

#[test]
fn short_left_shift_within_range() {
    let out = run_main("short s = 3; System.out.println((short)(s << 2));");
    assert_eq!(out, vec!["12"]);
}

#[test]
fn byte_right_shift_preserves_sign() {
    let out = run_main("byte b = -8; System.out.println((byte)(b >> 1));");
    assert_eq!(out, vec!["-4"]);
}

#[test]
fn byte_unsigned_right_shift_fills_zeros() {
    let out = run_main("byte b = -1; System.out.println((byte)(b >>> 1));");
    assert_eq!(out, vec!["127"]);
}

#[test]
fn byte_comparison_less_than() {
    let out = run_main("byte a = 2; byte b = 5; System.out.println(a < b);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn short_comparison_equal_values() {
    let out = run_main("short a = 40; short b = 40; System.out.println(a == b);");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn byte_array_length_and_first_element() {
    let out = run_main(
        "byte[] data = {1, 2, 3}; System.out.println(data.length); System.out.println(data[0]);",
    );
    assert_eq!(out, vec!["3", "1"]);
}

#[test]
fn short_array_sum_in_loop() {
    let out = run_main(
        "short[] vals = {10, 20, 30}; int sum = 0; for (short v : vals) { sum += v; } System.out.println(sum);",
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn byte_promoted_to_int_in_mixed_addition() {
    let out = run_main("byte b = 10; int n = 1000; System.out.println(b + n);");
    assert_eq!(out, vec!["1010"]);
}

#[test]
fn short_promoted_to_int_before_multiplication() {
    let out = run_main("short s = 50; int n = 3; System.out.println(s * n);");
    assert_eq!(out, vec!["150"]);
}

#[test]
fn byte_wrapper_value_of_and_unbox() {
    let out = run_main("Byte boxed = 8; byte raw = boxed; System.out.println(raw + 2);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn short_wrapper_value_of_and_unbox() {
    let out = run_main("Short boxed = 16; short raw = boxed; System.out.println(raw + 4);");
    assert_eq!(out, vec!["20"]);
}

#[test]
fn overload_prefers_byte_over_int() {
    let types = r#"
        static String pick(byte b) { return "byte"; }
        static String pick(int n) { return "int"; }
    "#;
    let out = run_in_main("System.out.println(pick((byte) 2));", types);
    assert_eq!(out, vec!["byte"]);
}
