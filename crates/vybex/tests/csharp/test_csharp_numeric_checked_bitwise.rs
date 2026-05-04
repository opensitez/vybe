use super::helpers::run_csharp;

macro_rules! csharp_case {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            assert_eq!(run_csharp($src), &[$($expected),*]);
        }
    };
}

csharp_case!(bitwise_and_keeps_shared_bits_only, r#"Console.WriteLine(6 & 3);"#, ["2"]);
csharp_case!(bitwise_or_combines_bits_from_both_operands, r#"Console.WriteLine(6 | 3);"#, ["7"]);
csharp_case!(bitwise_xor_flips_bits_that_differ, r#"Console.WriteLine(6 ^ 3);"#, ["5"]);
csharp_case!(bitwise_not_inverts_integer_bits, r#"Console.WriteLine(~5);"#, ["-6"]);
csharp_case!(left_shift_moves_bits_to_higher_positions, r#"Console.WriteLine(3 << 2);"#, ["12"]);
csharp_case!(right_shift_moves_bits_to_lower_positions, r#"Console.WriteLine(16 >> 3);"#, ["2"]);
csharp_case!(compound_shift_assignment_updates_variable_in_place, r#"int value = 5; value <<= 1; Console.WriteLine(value);"#, ["10"]);
csharp_case!(compound_bitwise_or_assignment_merges_mask_bits, r#"int value = 4; value |= 3; Console.WriteLine(value);"#, ["7"]);
csharp_case!(checked_block_throws_on_overflow_for_byte_addition, r#"try { checked { byte value = 255; value += 1; } Console.WriteLine("no-throw"); } catch (System.OverflowException) { Console.WriteLine("overflow"); }"#, ["overflow"]);
csharp_case!(unchecked_block_wraps_byte_overflow_without_throwing, r#"unchecked { byte value = 255; value += 1; Console.WriteLine(value); }"#, ["0"]);
csharp_case!(math_abs_returns_positive_magnitude, r#"Console.WriteLine(System.Math.Abs(-9));"#, ["9"]);
csharp_case!(math_min_selects_smaller_integer, r#"Console.WriteLine(System.Math.Min(4, 7));"#, ["4"]);
csharp_case!(math_max_selects_larger_integer, r#"Console.WriteLine(System.Math.Max(4, 7));"#, ["7"]);
csharp_case!(math_round_rounds_half_up_for_midpoint_value, r#"Console.WriteLine(System.Math.Round(4.5));"#, ["4"]);
csharp_case!(math_floor_truncates_toward_negative_infinity, r#"Console.WriteLine(System.Math.Floor(3.9));"#, ["3"]);
csharp_case!(math_ceiling_truncates_toward_positive_infinity, r#"Console.WriteLine(System.Math.Ceiling(3.1));"#, ["4"]);
csharp_case!(math_pow_raises_value_to_integer_exponent, r#"Console.WriteLine(System.Math.Pow(2, 5));"#, ["32"]);
csharp_case!(math_sqrt_returns_principal_square_root, r#"Console.WriteLine(System.Math.Sqrt(81));"#, ["9"]);
csharp_case!(decimal_addition_preserves_decimal_precision, r#"decimal left = 1.2m; decimal right = 2.3m; Console.WriteLine(left + right);"#, ["3.5"]);
csharp_case!(modulo_operator_returns_remainder_after_division, r#"Console.WriteLine(29 % 6);"#, ["5"]);