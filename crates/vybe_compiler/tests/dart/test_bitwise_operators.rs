//! Dart bitwise operators: &, |, ^, ~, <<, >>, compound assigns, precedence vs arithmetic.

dart_cases! {
    bitwise_and_two_positive_integers => {
        r#"void main() {
  print(12 & 10);
}"#,
        ["8"]
    };

    bitwise_and_with_zero_mask_clears_all_bits => {
        r#"void main() {
  print(0xFF & 0);
}"#,
        ["0"]
    };

    bitwise_and_with_full_mask_preserves_lower_bits => {
        r#"void main() {
  print(0xFF & 0x0F);
}"#,
        ["15"]
    };

    bitwise_or_combines_disjoint_bit_patterns => {
        r#"void main() {
  print(0xF0 | 0x0F);
}"#,
        ["255"]
    };

    bitwise_or_with_overlapping_bits_keeps_union => {
        r#"void main() {
  print(0xAA | 0x55);
}"#,
        ["255"]
    };

    bitwise_or_with_zero_operand_is_identity => {
        r#"void main() {
  print(42 | 0);
}"#,
        ["42"]
    };

    bitwise_xor_flips_matching_one_bits => {
        r#"void main() {
  print(0xFF ^ 0x0F);
}"#,
        ["240"]
    };

    bitwise_xor_with_self_yields_zero => {
        r#"void main() {
  print(73 ^ 73);
}"#,
        ["0"]
    };

    bitwise_xor_with_zero_preserves_operand => {
        r#"void main() {
  print(0x2A ^ 0);
}"#,
        ["42"]
    };

    bitwise_not_zero_is_negative_one => {
        r#"void main() {
  print(~0);
}"#,
        ["-1"]
    };

    bitwise_not_negative_one_is_zero => {
        r#"void main() {
  print(~(-1));
}"#,
        ["0"]
    };

    bitwise_not_positive_five => {
        r#"void main() {
  print(~5);
}"#,
        ["-6"]
    };

    double_bitwise_not_restores_positive_value => {
        r#"void main() {
  print(~~7);
}"#,
        ["7"]
    };

    left_shift_by_zero_preserves_value => {
        r#"void main() {
  print(99 << 0);
}"#,
        ["99"]
    };

    left_shift_by_one_doubles_value => {
        r#"void main() {
  print(3 << 1);
}"#,
        ["6"]
    };

    left_shift_by_four_multiplies_by_sixteen => {
        r#"void main() {
  print(1 << 4);
}"#,
        ["16"]
    };

    right_shift_by_zero_preserves_value => {
        r#"void main() {
  print(88 >> 0);
}"#,
        ["88"]
    };

    right_shift_divides_by_power_of_two => {
        r#"void main() {
  print(64 >> 2);
}"#,
        ["16"]
    };

    right_shift_large_value_by_three => {
        r#"void main() {
  print(256 >> 3);
}"#,
        ["32"]
    };

    left_shift_then_right_shift_truncates_high_bits => {
        r#"void main() {
  print((5 << 3) >> 1);
}"#,
        ["20"]
    };

    bitwise_and_assign_updates_variable => {
        r#"void main() {
  var x = 0xFF;
  x &= 0x0F;
  print(x);
}"#,
        ["15"]
    };

    bitwise_or_assign_merges_bit_patterns => {
        r#"void main() {
  var x = 0xF0;
  x |= 0x0F;
  print(x);
}"#,
        ["255"]
    };

    bitwise_xor_assign_toggles_selected_bits => {
        r#"void main() {
  var x = 0xFF;
  x ^= 0x0F;
  print(x);
}"#,
        ["240"]
    };

    left_shift_assign_multiplies_in_place => {
        r#"void main() {
  var x = 1;
  x <<= 3;
  print(x);
}"#,
        ["8"]
    };

    right_shift_assign_divides_in_place => {
        r#"void main() {
  var x = 32;
  x >>= 2;
  print(x);
}"#,
        ["8"]
    };

    compound_or_then_and_assign_sequence => {
        r#"void main() {
  var x = 0x10;
  x |= 0x01;
  x &= 0x11;
  print(x);
}"#,
        ["17"]
    };

    chained_xor_assign_returns_to_original => {
        r#"void main() {
  var x = 0xAB;
  x ^= 0xFF;
  x ^= 0xFF;
  print(x);
}"#,
        ["171"]
    };

    bitwise_and_has_higher_precedence_than_or => {
        r#"void main() {
  print(0xF0 | 0x0F & 0x05);
}"#,
        ["245"]
    };

    bitwise_xor_between_and_and_or => {
        r#"void main() {
  print(0xFF & 0xF0 ^ 0x0F);
}"#,
        ["255"]
    };

    addition_before_bitwise_and_masks_sum => {
        r#"void main() {
  print(2 + 3 & 4);
}"#,
        ["4"]
    };

    grouped_bitwise_and_before_addition => {
        r#"void main() {
  print(2 + (3 & 4));
}"#,
        ["2"]
    };

    left_shift_after_addition_in_operand => {
        r#"void main() {
  print(4 << 1 + 1);
}"#,
        ["16"]
    };

    addition_before_shift_then_left_shift => {
        r#"void main() {
  print(1 + 2 << 2);
}"#,
        ["12"]
    };

    multiplication_before_left_shift => {
        r#"void main() {
  print(2 * 3 << 1);
}"#,
        ["12"]
    };

    bitwise_and_with_negated_operand => {
        r#"void main() {
  print((-1) & 0xFF);
}"#,
        ["255"]
    };

    negative_operand_bitwise_and_preserves_low_bits => {
        r#"void main() {
  print(-8 & 7);
}"#,
        ["0"]
    };

    negative_right_shift_sign_extends => {
        r#"void main() {
  print(-8 >> 1);
}"#,
        ["-4"]
    };

    hex_literal_bitwise_and_decimal => {
        r#"void main() {
  print(0xABC & 0xFF);
}"#,
        ["188"]
    };

    hex_literal_bitwise_or_decimal => {
        r#"void main() {
  print(0x10 | 0x01);
}"#,
        ["17"]
    };

    nested_bitwise_and_or_expression => {
        r#"void main() {
  print((6 & 3) | (8 & 12));
}"#,
        ["10"]
    };

    bitwise_or_on_shifted_values => {
        r#"void main() {
  print((1 << 2) | (1 << 3));
}"#,
        ["12"]
    };

    right_shift_assign_on_negative_value => {
        r#"void main() {
  var x = -16;
  x >>= 2;
  print(x);
}"#,
        ["-4"]
    };

    xor_assign_clears_matching_bits_in_place => {
        r#"void main() {
  var x = 0b1010;
  x ^= 0b1100;
  print(x);
}"#,
        ["6"]
    };

    triple_compound_bitwise_sequence => {
        r#"void main() {
  var x = 0x0F;
  x <<= 1;
  x |= 0x01;
  x &= 0x1F;
  print(x);
}"#,
        ["31"]
    };

    unary_bitwise_not_before_and => {
        r#"void main() {
  print(~0 & 0xFF);
}"#,
        ["255"]
    };

    left_shift_on_negative_operand => {
        r#"void main() {
  print(-1 << 4);
}"#,
        ["-16"]
    };
}
