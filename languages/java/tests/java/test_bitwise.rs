use crate::helpers::run_main;

#[test]
fn bitwise_and_masks_common_bits() {
    let out = run_main("System.out.println(0b1100 & 0b1010);");
    assert_eq!(out, vec!["8"]);
}

#[test]
fn bitwise_and_with_zero_clears_all_bits() {
    let out = run_main("System.out.println(0b1111 & 0);");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn bitwise_and_with_ones_preserves_value() {
    let out = run_main("System.out.println(37 & 0xFF);");
    assert_eq!(out, vec!["37"]);
}

#[test]
fn bitwise_or_combines_set_bits() {
    let out = run_main("System.out.println(0b1100 | 0b1010);");
    assert_eq!(out, vec!["14"]);
}

#[test]
fn bitwise_or_with_zero_is_identity() {
    let out = run_main("System.out.println(25 | 0);");
    assert_eq!(out, vec!["25"]);
}

#[test]
fn bitwise_or_sets_high_bit() {
    let out = run_main("System.out.println(1 | 8);");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn bitwise_xor_sets_bits_different_in_operands() {
    let out = run_main("System.out.println(0b1100 ^ 0b1010);");
    assert_eq!(out, vec!["6"]);
}

#[test]
fn bitwise_xor_with_self_yields_zero() {
    let out = run_main("System.out.println(0b101101 ^ 0b101101);");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn bitwise_xor_toggles_single_bit() {
    let out = run_main("System.out.println(0b1000 ^ 0b0100);");
    assert_eq!(out, vec!["12"]);
}

#[test]
fn bitwise_not_inverts_all_bits() {
    let out = run_main("System.out.println(~0);");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn bitwise_not_of_positive_one_is_negative_two() {
    let out = run_main("System.out.println(~1);");
    assert_eq!(out, vec!["-2"]);
}

#[test]
fn bitwise_not_of_negative_one_is_zero() {
    let out = run_main("System.out.println(~(-1));");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn shift_left_multiplies_by_power_of_two() {
    let out = run_main("System.out.println(3 << 2);");
    assert_eq!(out, vec!["12"]);
}

#[test]
fn shift_left_by_zero_is_identity() {
    let out = run_main("System.out.println(99 << 0);");
    assert_eq!(out, vec!["99"]);
}

#[test]
fn shift_left_one_bit_doubles_value() {
    let out = run_main("System.out.println(5 << 1);");
    assert_eq!(out, vec!["10"]);
}

#[test]
fn signed_shift_right_preserves_sign_for_negative() {
    let out = run_main("System.out.println(-8 >> 1);");
    assert_eq!(out, vec!["-4"]);
}

#[test]
fn signed_shift_right_divides_positive_by_two() {
    let out = run_main("System.out.println(16 >> 2);");
    assert_eq!(out, vec!["4"]);
}

#[test]
fn signed_shift_right_by_zero_is_identity() {
    let out = run_main("System.out.println(7 >> 0);");
    assert_eq!(out, vec!["7"]);
}

#[test]
fn unsigned_shift_right_fills_with_zero() {
    let out = run_main("System.out.println(-1 >>> 1);");
    assert_eq!(out, vec!["2147483647"]);
}

#[test]
fn unsigned_shift_right_on_positive_matches_signed() {
    let out = run_main("System.out.println(32 >>> 2);");
    assert_eq!(out, vec!["8"]);
}

#[test]
fn unsigned_shift_right_large_positive_value() {
    let out = run_main("System.out.println(0x80000000 >>> 1);");
    assert_eq!(out, vec!["1073741824"]);
}

#[test]
fn bitwise_and_or_xor_combined_expression() {
    let out = run_main("System.out.println((6 & 3) | (5 ^ 1));");
    assert_eq!(out, vec!["6"]);
}

#[test]
fn bitwise_ops_on_variables() {
    let out = run_main(
        "int a = 0b1010; int b = 0b1100; System.out.println(a & b); System.out.println(a | b);",
    );
    assert_eq!(out, vec!["8", "14"]);
}

#[test]
fn shift_ops_on_variables() {
    let out = run_main("int n = 4; System.out.println(n << 1); System.out.println(n >> 1);");
    assert_eq!(out, vec!["8", "2"]);
}

#[test]
fn bitwise_not_on_variable() {
    let out = run_main("int x = 0; System.out.println(~x);");
    assert_eq!(out, vec!["-1"]);
}

#[test]
fn bitwise_and_with_hex_literals() {
    let out = run_main("System.out.println(0xF0 & 0x0F);");
    assert_eq!(out, vec!["0"]);
}

#[test]
fn bitwise_or_with_hex_literals() {
    let out = run_main("System.out.println(0xF0 | 0x0F);");
    assert_eq!(out, vec!["255"]);
}

#[test]
fn bitwise_xor_with_hex_literals() {
    let out = run_main("System.out.println(0xFF ^ 0x0F);");
    assert_eq!(out, vec!["240"]);
}

#[test]
fn unsigned_shift_right_by_zero_is_identity() {
    let out = run_main("System.out.println(-100 >>> 0);");
    assert_eq!(out, vec!["-100"]);
}

#[test]
fn chained_shifts_and_masks_extract_nibble() {
    let out = run_main("System.out.println((0xABCD & 0x00F0) >> 4);");
    assert_eq!(out, vec!["12"]);
}
