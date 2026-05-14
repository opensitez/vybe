use super::helpers::run_vb;

macro_rules! vb_expr_spec {
    ($name:ident, $expr:expr, $expected:expr) => {
        #[test]
        fn $name() {
            let src = format!(
                r#"Module M
    Sub Main()
        Console.WriteLine({})
    End Sub
End Module
"#,
                $expr
            );
            let out = run_vb(&src);
            assert_eq!(out, vec![$expected]);
        }
    };
}

macro_rules! vb_full_spec {
    ($name:ident, $src:expr, [$($expected:expr),* $(,)?]) => {
        #[test]
        fn $name() {
            let out = run_vb($src);
            assert_eq!(out, vec![$($expected),*]);
        }
    };
}

vb_expr_spec!(operator_spec_integer_division_truncates_positive_values, r#"7 \ 2"#, "3");
vb_expr_spec!(operator_spec_integer_division_truncates_negative_values, r#"-7 \ 2"#, "-3");
vb_expr_spec!(operator_spec_integer_division_returns_exact_quotient_when_evenly_divisible, r#"8 \ 2"#, "4");
vb_expr_spec!(operator_spec_floating_division_returns_fractional_result, r#"7 / 2"#, "3.5");
vb_expr_spec!(operator_spec_mod_returns_positive_remainder, r#"7 Mod 3"#, "1");
vb_expr_spec!(operator_spec_mod_preserves_sign_of_left_operand, r#"-7 Mod 3"#, "-1");
vb_expr_spec!(operator_spec_exponentiation_squares_number, r#"5 ^ 2"#, "25");
vb_expr_spec!(operator_spec_exponentiation_is_right_associative, r#"2 ^ 3 ^ 2"#, "512");
vb_expr_spec!(operator_spec_addition_combines_integers, r#"20 + 22"#, "42");
vb_expr_spec!(operator_spec_subtraction_reduces_integer, r#"20 - 8"#, "12");
vb_expr_spec!(operator_spec_multiplication_scales_integer, r#"6 * 7"#, "42");
vb_expr_spec!(operator_spec_unary_minus_negates_literal, r#"-5"#, "-5");
vb_expr_spec!(operator_spec_unary_plus_preserves_literal, r#"+5"#, "5");
vb_expr_spec!(operator_spec_not_inverts_boolean_true, r#"Not True"#, "false");
vb_expr_spec!(operator_spec_not_inverts_boolean_false, r#"Not False"#, "true");
vb_full_spec!(operator_spec_andalso_short_circuits_false_left_operand, r#"Module M
    Function Explode() As Boolean
        Console.WriteLine("boom")
        Return True
    End Function

    Sub Main()
        Console.WriteLine(False AndAlso Explode())
    End Sub
End Module"#, ["false"]);
vb_full_spec!(operator_spec_orelse_short_circuits_true_left_operand, r#"Module M
    Function Explode() As Boolean
        Console.WriteLine("boom")
        Return False
    End Function

    Sub Main()
        Console.WriteLine(True OrElse Explode())
    End Sub
End Module"#, ["true"]);
vb_expr_spec!(operator_spec_and_combines_boolean_truth, r#"True And False"#, "false");
vb_expr_spec!(operator_spec_or_combines_boolean_truth, r#"True Or False"#, "true");
vb_expr_spec!(operator_spec_xor_is_true_for_mixed_booleans, r#"True Xor False"#, "true");
vb_expr_spec!(operator_spec_xor_is_false_for_equal_booleans, r#"True Xor True"#, "false");
vb_expr_spec!(operator_spec_eqv_is_true_for_equal_booleans, r#"True Eqv True"#, "true");
vb_expr_spec!(operator_spec_eqv_is_false_for_different_booleans, r#"True Eqv False"#, "false");
vb_expr_spec!(operator_spec_imp_true_implies_false_is_false, r#"True Imp False"#, "false");
vb_expr_spec!(operator_spec_imp_false_implies_false_is_true, r#"False Imp False"#, "true");
vb_expr_spec!(operator_spec_less_than_compares_numbers, r#"3 < 5"#, "true");
vb_expr_spec!(operator_spec_greater_than_compares_numbers, r#"5 > 3"#, "true");
vb_expr_spec!(operator_spec_less_equal_accepts_equal_values, r#"5 <= 5"#, "true");
vb_expr_spec!(operator_spec_greater_equal_accepts_equal_values, r#"5 >= 5"#, "true");
vb_expr_spec!(operator_spec_not_equal_detects_difference, r#"5 <> 4"#, "true");
vb_expr_spec!(operator_spec_arithmetic_precedence_multiplies_before_addition, r#"2 + 3 * 4"#, "14");
vb_expr_spec!(operator_spec_parentheses_override_arithmetic_precedence, r#"(2 + 3) * 4"#, "20");
vb_expr_spec!(operator_spec_string_ampersand_concatenates_text_and_number, r#""count=" & 5"#, "count=5");
vb_expr_spec!(operator_spec_plus_adds_numbers_inside_parenthesized_expression, r#"(10 + 5) + 2"#, "17");
vb_expr_spec!(operator_spec_like_matches_wildcard_suffix, r#""visual" Like "vis*""#, "true");
vb_expr_spec!(operator_spec_like_matches_single_character_wildcard, r#""cat" Like "c?t""#, "true");
vb_expr_spec!(operator_spec_like_matches_digit_character_class, r#""A5" Like "A#""#, "true");
vb_expr_spec!(operator_spec_like_rejects_nonmatching_pattern, r#""dog" Like "c*""#, "false");
vb_expr_spec!(operator_spec_shift_left_moves_bits_up, r#"3 << 2"#, "12");
vb_expr_spec!(operator_spec_shift_right_moves_bits_down, r#"16 >> 2"#, "4");
vb_expr_spec!(operator_spec_boolean_precedence_applies_not_before_and, r#"Not False And False"#, "false");
vb_expr_spec!(operator_spec_boolean_precedence_applies_and_before_or, r#"True Or False And False"#, "true");
vb_expr_spec!(operator_spec_date_comparison_orders_earlier_before_later, r#"#5/14/2024# < #5/15/2024#"#, "true");
vb_expr_spec!(operator_spec_char_literal_compares_equal, r#""A"c = "A"c"#, "true");
vb_expr_spec!(operator_spec_string_compare_binary_orders_upper_before_lower, r#""A" < "a""#, "true");
vb_full_spec!(operator_spec_is_operator_matches_same_reference, r#"Class Box : End Class
Module M
    Sub Main()
        Dim value As New Box()
        Dim aliasValue As Box = value
        Console.WriteLine(value Is aliasValue)
    End Sub
End Module"#, ["true"]);
vb_full_spec!(operator_spec_isnot_operator_distinguishes_different_instances, r#"Class Box : End Class
Module M
    Sub Main()
        Dim left As New Box()
        Dim right As New Box()
        Console.WriteLine(left IsNot right)
    End Sub
End Module"#, ["true"]);
vb_full_spec!(operator_spec_nothing_isnot_object_reference, r#"Class Box : End Class
Module M
    Sub Main()
        Dim value As New Box()
        Console.WriteLine(Nothing IsNot value)
    End Sub
End Module"#, ["true"]);
vb_full_spec!(operator_spec_object_reference_is_self, r#"Class Box : End Class
Module M
    Sub Main()
        Dim value As New Box()
        Console.WriteLine(value Is value)
    End Sub
End Module"#, ["true"]);
vb_expr_spec!(operator_spec_concatenation_associates_left_to_right, r#""a" & 1 & 2"#, "a12");
