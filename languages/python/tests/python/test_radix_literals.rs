use crate::helpers::run_print;

#[test]
fn hex_literal_0xff() {
    assert_eq!(run_print("0xFF"), "255");
}

#[test]
fn hex_literal_lowercase() {
    assert_eq!(run_print("0xff"), "255");
}

#[test]
fn hex_literal_with_underscores() {
    assert_eq!(run_print("0xF_F"), "255");
}

#[test]
fn oct_literal_simple() {
    assert_eq!(run_print("0o10"), "8");
}

#[test]
fn oct_literal_with_underscores() {
    assert_eq!(run_print("0o1_0"), "8");
}

#[test]
fn bin_literal_simple() {
    assert_eq!(run_print("0b1010"), "10");
}

#[test]
fn bin_literal_with_underscores() {
    assert_eq!(run_print("0b1010_1010"), "170");
}

#[test]
fn decimal_with_underscore_separators() {
    assert_eq!(run_print("1_000_000"), "1000000");
}

#[test]
fn hex_builtin_positive() {
    assert_eq!(run_print("hex(255)"), "0xff");
}

#[test]
fn hex_builtin_zero() {
    assert_eq!(run_print("hex(0)"), "0x0");
}

#[test]
fn hex_builtin_negative() {
    assert_eq!(run_print("hex(-16)"), "-0x10");
}

#[test]
fn oct_builtin_positive() {
    assert_eq!(run_print("oct(8)"), "0o10");
}

#[test]
fn oct_builtin_zero() {
    assert_eq!(run_print("oct(0)"), "0o0");
}

#[test]
fn bin_builtin_positive() {
    assert_eq!(run_print("bin(10)"), "0b1010");
}

#[test]
fn bin_builtin_zero() {
    assert_eq!(run_print("bin(0)"), "0b0");
}

#[test]
fn int_from_hex_string() {
    assert_eq!(run_print("int('ff', 16)"), "255");
}

#[test]
fn int_from_oct_string() {
    assert_eq!(run_print("int('17', 8)"), "15");
}

#[test]
fn int_from_bin_string() {
    assert_eq!(run_print("int('1010', 2)"), "10");
}

#[test]
fn int_from_decimal_string() {
    assert_eq!(run_print("int('42')"), "42");
}

#[test]
fn int_from_float_truncates() {
    assert_eq!(run_print("int(3.9)"), "3");
}

#[test]
fn int_from_bool_true() {
    assert_eq!(run_print("int(True)"), "1");
}

#[test]
fn int_from_bool_false() {
    assert_eq!(run_print("int(False)"), "0");
}

#[test]
fn float_from_int_literal() {
    assert_eq!(run_print("float(7)"), "7.0");
}

#[test]
fn float_from_string() {
    assert_eq!(run_print("float('3.14')"), "3.14");
}

#[test]
fn float_scientific_notation() {
    assert_eq!(run_print("1e3"), "1000.0");
}

#[test]
fn float_negative_scientific() {
    assert_eq!(run_print("1e-2"), "0.01");
}

#[test]
fn hex_addition_with_decimal() {
    assert_eq!(run_print("0x10 + 16"), "32");
}

#[test]
fn bin_addition_with_decimal() {
    assert_eq!(run_print("0b11 + 1"), "4");
}

#[test]
fn oct_multiplication() {
    assert_eq!(run_print("0o10 * 2"), "16");
}

#[test]
fn underscore_in_float_literal() {
    assert_eq!(run_print("3_14.15_9"), "314.159");
}

#[test]
fn hex_literal_large() {
    assert_eq!(run_print("0xDEAD"), "57005");
}

#[test]
fn bin_literal_eight_bits() {
    assert_eq!(run_print("0b11111111"), "255");
}

#[test]
fn oct_literal_max_single_digit() {
    assert_eq!(run_print("0o7"), "7");
}

#[test]
fn int_base36_lowercase() {
    assert_eq!(run_print("int('z', 36)"), "35");
}

#[test]
fn int_base2_with_prefix_stripped() {
    assert_eq!(run_print("int('1010', 2)"), "10");
}

#[test]
fn hex_of_power_of_two() {
    assert_eq!(run_print("hex(1024)"), "0x400");
}

#[test]
fn bin_of_five() {
    assert_eq!(run_print("bin(5)"), "0b101");
}

#[test]
fn oct_of_sixty_four() {
    assert_eq!(run_print("oct(64)"), "0o100");
}

#[test]
fn literal_mix_in_expression() {
    assert_eq!(run_print("0xF + 0b1 + 0o0"), "16");
}

#[test]
fn negative_hex_literal() {
    assert_eq!(run_print("-0x10"), "-16");
}

#[test]
fn negative_bin_literal() {
    assert_eq!(run_print("-0b101"), "-5");
}

#[test]
fn chr_of_sixty_five_is_a() {
    assert_eq!(run_print("chr(65)"), "A");
}

#[test]
fn ord_of_uppercase_a() {
    assert_eq!(run_print("ord('A')"), "65");
}

#[test]
fn ord_of_digit_zero() {
    assert_eq!(run_print("ord('0')"), "48");
}

#[test]
fn chr_of_newline() {
    assert_eq!(run_print("repr(chr(10))"), "'\\n'");
}

#[test]
fn int_roundtrip_hex_string() {
    assert_eq!(run_print("int(hex(200), 0)"), "200");
}

#[test]
fn float_whole_number_has_point_zero() {
    assert_eq!(run_print("float(0)"), "0.0");
}

#[test]
fn hex_literal_in_comparison() {
    assert_eq!(run_print("0x10 > 15"), "True");
}

#[test]
fn bin_literal_equality_with_decimal() {
    assert_eq!(run_print("0b1010 == 10"), "True");
}

#[test]
fn oct_literal_less_than_hex_same_value() {
    assert_eq!(run_print("0o10 == 0x8"), "True");
}
