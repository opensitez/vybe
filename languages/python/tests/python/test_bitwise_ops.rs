use crate::helpers::{run_print, run_python_one};

#[test]
fn bitwise_and() {
    assert_eq!(run_print("5 & 3"), "1");
}

#[test]
fn bitwise_or() {
    assert_eq!(run_print("5 | 2"), "7");
}

#[test]
fn bitwise_xor() {
    assert_eq!(run_print("5 ^ 3"), "6");
}

#[test]
fn bitwise_not() {
    assert_eq!(run_print("~0"), "-1");
}

#[test]
fn bitwise_left_shift() {
    assert_eq!(run_print("1 << 4"), "16");
}

#[test]
fn bitwise_right_shift() {
    assert_eq!(run_print("16 >> 2"), "4");
}

#[test]
fn bitwise_and_with_zero() {
    assert_eq!(run_print("255 & 0"), "0");
}

#[test]
fn bitwise_or_with_zero() {
    assert_eq!(run_print("10 | 0"), "10");
}

#[test]
fn bitwise_xor_self_zero() {
    assert_eq!(run_print("7 ^ 7"), "0");
}

#[test]
fn bitwise_shift_zero_bits() {
    assert_eq!(run_print("99 << 0"), "99");
}

#[test]
fn bitwise_complex_masking() {
    assert_eq!(run_print("0b11110000 & 0b10101010"), "160");
}

#[test]
fn bitwise_set_flag() {
    assert_eq!(run_print("flags | 0b100"), "4");
}

#[test]
fn bitwise_clear_flag() {
    assert_eq!(run_print("7 & ~2"), "5");
}

#[test]
fn bitwise_toggle_flag() {
    assert_eq!(run_print("5 ^ 1"), "4");
}

#[test]
fn bitwise_in_place_and() {
    assert_eq!(run_python_one("x = 15\nx &= 10\nprint(x)\n"), "10");
}

#[test]
fn bitwise_in_place_or() {
    assert_eq!(run_python_one("x = 1\nx |= 2\nprint(x)\n"), "3");
}

#[test]
fn bitwise_in_place_xor() {
    assert_eq!(run_python_one("x = 5\nx ^= 3\nprint(x)\n"), "6");
}

#[test]
fn bitwise_in_place_lshift() {
    assert_eq!(run_python_one("x = 2\nx <<= 3\nprint(x)\n"), "16");
}

#[test]
fn bitwise_in_place_rshift() {
    assert_eq!(run_python_one("x = 32\nx >>= 2\nprint(x)\n"), "8");
}

#[test]
fn bitwise_negative_right_shift() {
    assert_eq!(run_print("-8 >> 1"), "-4");
}

#[test]
fn bitwise_large_shift() {
    assert_eq!(run_print("1 << 10"), "1024");
}

#[test]
fn bitwise_chained_or_and() {
    assert_eq!(run_print("(5 | 3) & 6"), "6");
}

#[test]
fn bitwise_precedence_over_comparison() {
    assert_eq!(run_print("(1 << 2) > 3"), "True");
}

#[test]
fn bitwise_bool_and_still_logical() {
    assert_eq!(run_print("True & True"), "True");
}

#[test]
fn bitwise_bool_or_still_logical() {
    assert_eq!(run_print("False | True"), "True");
}

#[test]
fn bitwise_bool_xor() {
    assert_eq!(run_print("True ^ False"), "True");
}

#[test]
fn bitwise_int_from_hex() {
    assert_eq!(run_print("0xF & 0x3"), "3");
}

#[test]
fn bitwise_int_from_binary() {
    assert_eq!(run_print("0b1010 | 0b0101"), "15");
}

#[test]
fn bitwise_count_bits_manual() {
    assert_eq!(
        run_python_one("n = 7\ncount = 0\nwhile n:\n count += n & 1\n n >>= 1\nprint(count)\n"),
        "3"
    );
}

#[test]
fn bitwise_parity_check() {
    assert_eq!(run_print("13 & 1"), "1");
}

#[test]
fn bitwise_even_check_clear_last_bit() {
    assert_eq!(run_print("10 & ~1"), "10");
}

#[test]
fn bitwise_swap_xor_trick() {
    assert_eq!(
        run_python_one("a, b = 5, 7\na ^= b\nb ^= a\na ^= b\nprint(a, b)\n"),
        "7 5"
    );
}

#[test]
fn bitwise_mask_low_nibble() {
    assert_eq!(run_print("0xAB & 0x0F"), "11");
}

#[test]
fn bitwise_mask_high_nibble() {
    assert_eq!(run_print("(0xAB & 0xF0) >> 4"), "10");
}

#[test]
fn bitwise_combine_permissions() {
    assert_eq!(run_print("0b001 | 0b010 | 0b100"), "7");
}

#[test]
fn bitwise_test_bit_set() {
    assert_eq!(run_print("(8 >> 3) & 1"), "1");
}

#[test]
fn bitwise_test_bit_clear() {
    assert_eq!(run_print("(4 >> 2) & 1"), "1");
}

#[test]
fn bitwise_shift_then_mask() {
    assert_eq!(run_print("(0xFF00 >> 8) & 0xFF"), "255");
}

#[test]
fn bitwise_not_small_number() {
    assert_eq!(run_print("~~5"), "5");
}

#[test]
fn bitwise_or_assign_build_flags() {
    assert_eq!(run_python_one("f = 0\nf |= 1\nf |= 4\nprint(f)\n"), "5");
}

#[test]
fn bitwise_and_assign_keep_lower_two_bits() {
    assert_eq!(run_python_one("x = 0b110101\nx &= 0b11\nprint(x)\n"), "1");
}

#[test]
fn bitwise_xor_assign_toggle_twice() {
    assert_eq!(run_python_one("x = 9\nx ^= 5\nx ^= 5\nprint(x)\n"), "9");
}

#[test]
fn bitwise_on_negative_and_positive() {
    assert_eq!(run_print("(-1) & 7"), "7");
}

#[test]
fn bitwise_shift_bounds_practical() {
    assert_eq!(run_print("(1 << 8) - 1"), "255");
}

#[test]
fn bitwise_multi_or_chain() {
    assert_eq!(run_print("1 | 2 | 4 | 8"), "15");
}
