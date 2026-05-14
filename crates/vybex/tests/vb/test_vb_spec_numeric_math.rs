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

vb_expr_spec!(numeric_spec_abs_negative_integer_becomes_positive, r#"Abs(-7)"#, "7");
vb_expr_spec!(numeric_spec_abs_positive_integer_stays_positive, r#"Abs(7)"#, "7");
vb_expr_spec!(numeric_spec_sqr_returns_integer_root_for_perfect_square, r#"Sqr(81)"#, "9");
vb_expr_spec!(numeric_spec_sqr_returns_fractional_root_for_decimal_value, r#"Sqr(2.25)"#, "1.5");
vb_expr_spec!(numeric_spec_sin_of_zero_is_zero, r#"Sin(0)"#, "0");
vb_expr_spec!(numeric_spec_cos_of_zero_is_one, r#"Cos(0)"#, "1");
vb_expr_spec!(numeric_spec_tan_of_zero_is_zero, r#"Tan(0)"#, "0");
vb_expr_spec!(numeric_spec_exp_of_zero_is_one, r#"Exp(0)"#, "1");
vb_expr_spec!(numeric_spec_log_of_one_is_zero, r#"Log(1)"#, "0");
vb_expr_spec!(numeric_spec_atn_of_zero_is_zero, r#"Atn(0)"#, "0");
vb_expr_spec!(numeric_spec_round_half_value_can_round_to_even_integer, r#"Round(1.5)"#, "2");
vb_expr_spec!(numeric_spec_round_uses_bankers_rounding_for_even_target, r#"Round(2.5)"#, "2");
vb_expr_spec!(numeric_spec_round_with_digits_preserves_requested_scale, r#"Round(2.55, 1)"#, "2.6");
vb_expr_spec!(numeric_spec_cint_rounds_half_to_even_integer, r#"CInt(4.5)"#, "4");
vb_expr_spec!(numeric_spec_cint_rounds_above_half_upward, r#"CInt(4.6)"#, "5");
vb_expr_spec!(numeric_spec_math_abs_handles_negative_decimal, r#"Math.Abs(-7.25)"#, "7.25");
vb_expr_spec!(numeric_spec_math_max_returns_larger_value, r#"Math.Max(3, 9)"#, "9");
vb_expr_spec!(numeric_spec_math_min_returns_smaller_value, r#"Math.Min(3, 9)"#, "3");
vb_expr_spec!(numeric_spec_math_pow_raises_base_to_exponent, r#"Math.Pow(3, 4)"#, "81");
vb_expr_spec!(numeric_spec_math_sqrt_returns_expected_root, r#"Math.Sqrt(144)"#, "12");
vb_expr_spec!(numeric_spec_math_round_without_digits_uses_default_precision, r#"Math.Round(12.5)"#, "12");
vb_expr_spec!(numeric_spec_math_round_with_digits_returns_decimal_string, r#"Math.Round(12.345, 2)"#, "12.35");
vb_expr_spec!(numeric_spec_math_floor_drops_fractional_part, r#"Math.Floor(12.9)"#, "12");
vb_expr_spec!(numeric_spec_math_ceiling_advances_fractional_part, r#"Math.Ceiling(12.1)"#, "13");
vb_expr_spec!(numeric_spec_math_truncate_removes_fraction_without_rounding, r#"Math.Truncate(12.9)"#, "12");
vb_expr_spec!(numeric_spec_math_sign_returns_negative_indicator, r#"Math.Sign(-12)"#, "-1");
vb_expr_spec!(numeric_spec_math_sign_returns_zero_indicator, r#"Math.Sign(0)"#, "0");
vb_expr_spec!(numeric_spec_math_sign_returns_positive_indicator, r#"Math.Sign(12)"#, "1");
vb_expr_spec!(numeric_spec_math_pi_can_be_rounded_to_two_decimals, r#"Round(Math.PI, 2)"#, "3.14");
vb_expr_spec!(numeric_spec_math_e_can_be_rounded_to_two_decimals, r#"Round(Math.E, 2)"#, "2.72");
vb_expr_spec!(numeric_spec_math_sin_of_pi_over_two_rounds_to_one, r#"Round(Math.Sin(Math.PI / 2), 6)"#, "1");
vb_expr_spec!(numeric_spec_math_cos_of_pi_rounds_to_negative_one, r#"Round(Math.Cos(Math.PI), 6)"#, "-1");
vb_expr_spec!(numeric_spec_math_tan_of_pi_over_four_rounds_to_one, r#"Round(Math.Tan(Math.PI / 4), 6)"#, "1");
vb_expr_spec!(numeric_spec_math_log10_of_thousand_is_three, r#"Math.Log10(1000)"#, "3");
vb_expr_spec!(numeric_spec_math_exp_of_one_rounds_to_e, r#"Round(Math.Exp(1), 2)"#, "2.72");
vb_expr_spec!(numeric_spec_math_atan_of_one_rounds_to_pi_over_four, r#"Round(Math.Atan(1), 6)"#, "0.785398");
vb_expr_spec!(numeric_spec_math_pow_zero_exponent_returns_one, r#"Math.Pow(9, 0)"#, "1");
vb_expr_spec!(numeric_spec_math_sqrt_of_zero_returns_zero, r#"Math.Sqrt(0)"#, "0");
vb_expr_spec!(numeric_spec_round_negative_half_value_can_round_to_even_integer, r#"Round(-1.5)"#, "-2");
vb_expr_spec!(numeric_spec_cdec_value_can_be_rounded_to_two_places, r#"Round(CDec("12.345"), 2)"#, "12.34");
vb_expr_spec!(numeric_spec_math_max_works_for_negative_numbers, r#"Math.Max(-3, -9)"#, "-3");
vb_expr_spec!(numeric_spec_math_min_works_for_negative_numbers, r#"Math.Min(-3, -9)"#, "-9");
vb_expr_spec!(numeric_spec_math_abs_of_zero_is_zero, r#"Math.Abs(0)"#, "0");
vb_expr_spec!(numeric_spec_math_pow_even_exponent_makes_negative_base_positive, r#"Math.Pow(-3, 2)"#, "9");
vb_expr_spec!(numeric_spec_math_pow_odd_exponent_keeps_negative_base_negative, r#"Math.Pow(-3, 3)"#, "-27");
vb_expr_spec!(numeric_spec_sqr_of_one_returns_one, r#"Sqr(1)"#, "1");
vb_expr_spec!(numeric_spec_cdbl_text_can_be_rounded_to_two_places, r#"Round(CDbl("2.345"), 2)"#, "2.35");
vb_expr_spec!(numeric_spec_math_floor_of_negative_fraction_goes_more_negative, r#"Math.Floor(-2.1)"#, "-3");
vb_expr_spec!(numeric_spec_math_ceiling_of_negative_fraction_moves_toward_zero, r#"Math.Ceiling(-2.9)"#, "-2");
vb_expr_spec!(numeric_spec_math_truncate_of_negative_fraction_moves_toward_zero, r#"Math.Truncate(-2.9)"#, "-2");
