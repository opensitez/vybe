//! Bitwise operators `&`, `|`, `^`, `~`, `<<`, `>>`, and `>>>` on integer types.
use super::helpers::run_csharp;

#[test]
fn bitwise_and_masks_bits() {
    assert_eq!(run_csharp(r#"Console.WriteLine(0b1100 & 0b1010);"#), &["8"]);
}

#[test]
fn bitwise_or_sets_bits() {
    assert_eq!(
        run_csharp(r#"Console.WriteLine(0b1100 | 0b0011);"#),
        &["15"]
    );
}

#[test]
fn bitwise_xor_toggles_differing_bits() {
    assert_eq!(run_csharp(r#"Console.WriteLine(0b1010 ^ 0b1100);"#), &["6"]);
}

#[test]
fn bitwise_not_inverts_all_bits_of_byte() {
    assert_eq!(
        run_csharp(r#"byte b = 0b11110000; Console.WriteLine((byte)(~b));"#),
        &["15"]
    );
}

#[test]
fn left_shift_multiplies_by_power_of_two() {
    assert_eq!(run_csharp(r#"Console.WriteLine(1 << 4);"#), &["16"]);
}

#[test]
fn right_shift_divides_by_power_of_two() {
    assert_eq!(run_csharp(r#"Console.WriteLine(64 >> 3);"#), &["8"]);
}

#[test]
fn signed_right_shift_preserves_sign_bit_for_negative() {
    assert_eq!(run_csharp(r#"Console.WriteLine(-8 >> 1);"#), &["-4"]);
}

#[test]
fn compound_bitwise_and_assign_updates_in_place() {
    assert_eq!(
        run_csharp(r#"int x = 0b1111; x &= 0b0101; Console.WriteLine(x);"#),
        &["5"]
    );
}

#[test]
fn compound_or_assign_sets_bits_in_place() {
    assert_eq!(
        run_csharp(r#"int x = 0b1000; x |= 0b0011; Console.WriteLine(x);"#),
        &["11"]
    );
}

#[test]
fn bit_test_using_and_with_power_of_two_mask() {
    assert_eq!(
        run_csharp(r#"int flags = 0b1010; Console.WriteLine((flags & 0b0010) != 0);"#),
        &["True"]
    );
}
