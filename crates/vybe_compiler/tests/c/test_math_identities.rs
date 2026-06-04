use super::helpers::*;

macro_rules! c_cases {
    ($($name:ident => { body: $body:expr, expect: [$($expected:expr),* $(,)?] }),* $(,)?) => {
        $(
            #[test]
            fn $name() {
                assert_program(&["<stdio.h>", "<math.h>"], "", $body, &[$($expected),*]);
            }
        )*
    };
}

c_cases! {
    sin_squared_plus_cos_squared_is_one_at_zero => { body: "double x = 0.0; printf(\"%.3f\\n\", sin(x) * sin(x) + cos(x) * cos(x)); return 0;", expect: ["1.000"] },
    sin_squared_plus_cos_squared_is_one_at_quarter => { body: "double x = 0.25; printf(\"%.3f\\n\", sin(x) * sin(x) + cos(x) * cos(x)); return 0;", expect: ["1.000"] },
    exp_log_round_trip_on_positive_number => { body: "printf(\"%.3f\\n\", exp(log(7.0))); return 0;", expect: ["7.000"] },
    log10_pow_round_trip_on_power_of_ten => { body: "printf(\"%.3f\\n\", log10(pow(10.0, 4.0))); return 0;", expect: ["4.000"] },
    sqrt_pow_round_trip_on_square_value => { body: "printf(\"%.3f\\n\", sqrt(pow(5.0, 2.0))); return 0;", expect: ["5.000"] },
    pow_of_zero_exponent_is_one => { body: "printf(\"%.1f\\n\", pow(9.0, 0.0)); return 0;", expect: ["1.0"] },
    pow_of_one_base_is_one => { body: "printf(\"%.1f\\n\", pow(1.0, 99.0)); return 0;", expect: ["1.0"] },
    fmod_of_larger_divisor_returns_dividend => { body: "printf(\"%.1f\\n\", fmod(3.0, 7.0)); return 0;", expect: ["3.0"] },
    floor_of_integer_like_value_is_same_value => { body: "printf(\"%.1f\\n\", floor(5.0)); return 0;", expect: ["5.0"] },
    ceil_of_integer_like_value_is_same_value => { body: "printf(\"%.1f\\n\", ceil(5.0)); return 0;", expect: ["5.0"] },
    round_of_integer_like_value_is_same_value => { body: "printf(\"%.1f\\n\", round(5.0)); return 0;", expect: ["5.0"] },
    fabs_of_positive_value_is_unchanged => { body: "printf(\"%.1f\\n\", fabs(5.0)); return 0;", expect: ["5.0"] },
    atan2_zero_one_is_zero => { body: "printf(\"%.3f\\n\", atan2(0.0, 1.0)); return 0;", expect: ["0.000"] },
    atan2_zero_negative_one_is_pi => { body: "printf(\"%.3f\\n\", atan2(0.0, -1.0)); return 0;", expect: ["3.142"] },
    cos_of_two_pi_is_one => { body: "printf(\"%.3f\\n\", cos(2.0 * M_PI)); return 0;", expect: ["1.000"] },
    sin_of_two_pi_is_zero => { body: "printf(\"%.3f\\n\", sin(2.0 * M_PI)); return 0;", expect: ["0.000"] },
    tan_of_negative_quarter_pi_is_negative_one => { body: "printf(\"%.3f\\n\", tan(-M_PI / 4.0)); return 0;", expect: ["-1.000"] },
    asin_sin_round_trip_small_angle => { body: "double x = 0.5; printf(\"%.3f\\n\", asin(sin(x))); return 0;", expect: ["0.500"] },
    acos_cos_round_trip_small_angle => { body: "double x = 0.5; printf(\"%.3f\\n\", acos(cos(x))); return 0;", expect: ["0.500"] },
    log_of_exp_round_trip_small_number => { body: "double x = 0.5; printf(\"%.3f\\n\", log(exp(x))); return 0;", expect: ["0.500"] },
    exp_of_log_round_trip_fraction => { body: "printf(\"%.3f\\n\", exp(log(0.5))); return 0;", expect: ["0.500"] },
    pow_and_fmod_can_be_composed => { body: "printf(\"%.1f\\n\", fmod(pow(3.0, 3.0), 5.0)); return 0;", expect: ["2.0"] },
    ceil_of_negative_small_fraction_is_zero => { body: "printf(\"%.0f\\n\", ceil(-0.2)); return 0;", expect: ["0"] },
    floor_of_negative_small_fraction_is_minus_one => { body: "printf(\"%.0f\\n\", floor(-0.2)); return 0;", expect: ["-1"] },
    round_of_negative_small_fraction_is_zero => { body: "printf(\"%.0f\\n\", round(-0.2)); return 0;", expect: ["0"] },
    log10_of_hundred_is_two => { body: "printf(\"%.3f\\n\", log10(100.0)); return 0;", expect: ["2.000"] },
    log_of_square_rooted_e_squared_is_one => { body: "printf(\"%.3f\\n\", log(sqrt(M_E * M_E))); return 0;", expect: ["1.000"] },
    pow_of_fractional_base_and_integer_exponent_preserves_fraction => { body: "printf(\"%.3f\\n\", pow(0.5, 2.0)); return 0;", expect: ["0.250"] },
    sqrt_of_two_matches_constant_relation => { body: "printf(\"%.3f\\n\", sqrt(2.0)); return 0;", expect: ["1.414"] },
    fabs_of_nested_negative_expression_is_positive => { body: "printf(\"%.1f\\n\", fabs(-3.0 * 2.0)); return 0;", expect: ["6.0"] },
    atan_of_tan_round_trip_near_zero => { body: "double x = 0.2; printf(\"%.3f\\n\", atan(tan(x))); return 0;", expect: ["0.200"] },
    exp_zero_log_one_identity_chain_is_one => { body: "printf(\"%.3f\\n\", exp(log(1.0))); return 0;", expect: ["1.000"] },
    pow_can_recreate_multiplication_result => { body: "printf(\"%.0f\\n\", pow(2.0, 5.0)); return 0;", expect: ["32"] },
    fmod_of_exact_multiple_is_zero => { body: "printf(\"%.1f\\n\", fmod(20.0, 4.0)); return 0;", expect: ["0.0"] },
    trigonometric_expression_can_feed_comparison => { body: "printf(\"%d\\n\", sin(M_PI / 2.0) > 0.9); return 0;", expect: ["1"] },
    logarithm_of_e_squared_is_two => { body: "printf(\"%.3f\\n\", log(M_E * M_E)); return 0;", expect: ["2.000"] },
    cosine_even_function_matches_positive_and_negative_inputs => { body: "printf(\"%d\\n\", cos(0.5) == cos(-0.5)); return 0;", expect: ["1"] },
    sine_odd_function_negates_positive_input => { body: "printf(\"%.3f\\n\", sin(-0.5) + sin(0.5)); return 0;", expect: ["0.000"] },
    floor_plus_ceil_of_same_fraction_brackets_value => { body: "printf(\"%.0f %.0f\\n\", floor(3.2), ceil(3.2)); return 0;", expect: ["3 4"] },
    exp_and_log_of_two_can_round_trip => { body: "printf(\"%.3f\\n\", exp(log(2.0))); return 0;", expect: ["2.000"] }
}
