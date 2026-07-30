//! dart:math via the `math` namespace: min/max, pow, sqrt, trig, constants, abs.

dart_cases! {
    math_sqrt_of_perfect_square => {
        r#"void main() {
  print(math.sqrt(81));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['9.0'].
        ["9.0"]
    };

    math_sqrt_of_zero => {
        r#"void main() {
  print(math.sqrt(0));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['0.0'].
        ["0.0"]
    };

    math_sqrt_of_one => {
        r#"void main() {
  print(math.sqrt(1));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['1.0'].
        ["1.0"]
    };

    math_sqrt_of_four => {
        r#"void main() {
  print(math.sqrt(4));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['2.0'].
        ["2.0"]
    };

    math_sqrt_for_pythagorean_triple => {
        r#"void main() {
  print(math.sqrt(3 * 3 + 4 * 4));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['5.0'].
        ["5.0"]
    };

    math_pow_squares_small_integer => {
        r#"void main() {
  print(math.pow(5, 2));
}"#,
        ["25"]
    };

    math_pow_cubes_three => {
        r#"void main() {
  print(math.pow(3, 3));
}"#,
        ["27"]
    };

    math_pow_zero_exponent => {
        r#"void main() {
  print(math.pow(9, 0));
}"#,
        ["1"]
    };

    math_pow_one_exponent => {
        r#"void main() {
  print(math.pow(7, 1));
}"#,
        ["7"]
    };

    math_pow_two_to_ten => {
        r#"void main() {
  print(math.pow(2, 10));
}"#,
        ["1024"]
    };

    math_pow_negative_exponent_fraction => {
        r#"void main() {
  print(math.pow(2, -1));
}"#,
        ["0.5"]
    };

    math_min_picks_smaller_of_two => {
        r#"void main() {
  print(math.min(4, 9));
}"#,
        ["4"]
    };

    math_min_with_negative_operand => {
        r#"void main() {
  print(math.min(-3, 2));
}"#,
        ["-3"]
    };

    math_min_equal_arguments => {
        r#"void main() {
  print(math.min(6, 6));
}"#,
        ["6"]
    };

    math_max_picks_larger_of_two => {
        r#"void main() {
  print(math.max(4, 9));
}"#,
        ["9"]
    };

    math_max_with_negative_operand => {
        r#"void main() {
  print(math.max(-3, 2));
}"#,
        ["2"]
    };

    math_max_equal_arguments => {
        r#"void main() {
  print(math.max(6, 6));
}"#,
        ["6"]
    };

    math_sin_of_zero => {
        r#"void main() {
  print(math.sin(0));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['0.0'].
        ["0.0"]
    };

    math_cos_of_zero => {
        r#"void main() {
  print(math.cos(0));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['1.0'].
        ["1.0"]
    };

    math_sin_of_pi_over_six => {
        r#"void main() {
  print(math.sin(math.pi / 6));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['0.49999999999999994'].
        ["0.49999999999999994"]
    };

    math_cos_of_pi_over_three => {
        r#"void main() {
  print(math.cos(math.pi / 3));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['0.5000000000000001'].
        ["0.5000000000000001"]
    };

    math_sin_of_pi_over_two_near_one => {
        r#"void main() {
  print(math.sin(math.pi / 2) > 0.99);
}"#,
        ["true"]
    };

    math_cos_of_pi_near_negative_one => {
        r#"void main() {
  print(math.cos(math.pi) < -0.99);
}"#,
        ["true"]
    };

    math_tan_of_zero => {
        r#"void main() {
  print(math.tan(0));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['0.0'].
        ["0.0"]
    };

    math_asin_of_zero => {
        r#"void main() {
  print(math.asin(0));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['0.0'].
        ["0.0"]
    };

    math_acos_of_one => {
        r#"void main() {
  print(math.acos(1));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['0.0'].
        ["0.0"]
    };

    math_atan_of_zero => {
        r#"void main() {
  print(math.atan(0));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['0.0'].
        ["0.0"]
    };

    math_atan2_first_quadrant => {
        r#"void main() {
  print(math.atan2(1, 1) > 0);
}"#,
        ["true"]
    };

    math_atan2_second_quadrant => {
        r#"void main() {
  print(math.atan2(1, -1) > 1);
}"#,
        ["true"]
    };

    math_exp_of_zero => {
        r#"void main() {
  print(math.exp(0));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['1.0'].
        ["1.0"]
    };

    math_log_of_one => {
        r#"void main() {
  print(math.log(1));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['0.0'].
        ["0.0"]
    };

    math_log_of_e_is_one => {
        r#"void main() {
  print(math.log(math.e) > 0.999);
}"#,
        ["true"]
    };

    math_pi_between_three_and_four => {
        r#"void main() {
  print(math.pi > 3 && math.pi < 4);
}"#,
        ["true"]
    };

    math_e_between_two_and_three => {
        r#"void main() {
  print(math.e > 2 && math.e < 3);
}"#,
        ["true"]
    };

    math_pi_greater_than_three_point_one_four => {
        r#"void main() {
  print(math.pi > 3.14);
}"#,
        ["true"]
    };

    math_e_greater_than_two_point_seven => {
        r#"void main() {
  print(math.e > 2.7);
}"#,
        ["true"]
    };

    math_sqrt2_constant_near_root_two => {
        r#"void main() {
  print(math.sqrt2 > 1.414 && math.sqrt2 < 1.415);
}"#,
        ["true"]
    };

    math_sqrt_matches_sqrt2_constant => {
        r#"void main() {
  print((math.sqrt(2) - math.sqrt2).abs() < 0.0001);
}"#,
        ["true"]
    };

    math_ln2_positive_less_than_one => {
        r#"void main() {
  print(math.ln2 > 0 && math.ln2 < 1);
}"#,
        ["true"]
    };

    math_ln10_between_two_and_three => {
        r#"void main() {
  print(math.ln10 > 2 && math.ln10 < 3);
}"#,
        ["true"]
    };

    double_abs_on_negative_fraction => {
        r#"void main() {
  print((-9.75).abs());
}"#,
        ["9.75"]
    };

    double_abs_on_positive_fraction => {
        r#"void main() {
  print(3.25.abs());
}"#,
        ["3.25"]
    };

    double_abs_on_negative_integer_as_double => {
        r#"void main() {
  print((-42.0).abs());
}"#,
        ["42.0"]
    };

    double_abs_after_math_sqrt => {
        r#"void main() {
  print((-math.sqrt(16)).abs());
}"#,
        ["4.0"]
    };

    math_min_clamps_negative_temperature_pair => {
        r#"void main() {
  print(math.min(-5.0, -12.0));
}"#,
        ["-12.0"]
    };

    math_max_clamps_positive_temperature_pair => {
        r#"void main() {
  print(math.max(18.5, 22.0));
}"#,
        ["22.0"]
    };

    math_pow_composes_with_sqrt => {
        r#"void main() {
  print(math.sqrt(math.pow(6, 2)));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['6.0'].
        ["6.0"]
    };

    math_sin_cos_pythagorean_on_unit_circle => {
        r#"void main() {
  var s = math.sin(math.pi / 4);
  var c = math.cos(math.pi / 4);
  print((s * s + c * c) > 0.99);
}"#,
        ["true"]
    };

    math_sqrt_of_pow_squared => {
        r#"void main() {
  print(math.sqrt(math.pow(11, 2)));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['11.0'].
        ["11.0"]
    };

    math_pow_then_min_with_smaller_value => {
        r#"void main() {
  print(math.min(math.pow(2, 5), 40));
}"#,
        ["32"]
    };

    math_pow_then_max_with_larger_value => {
        r#"void main() {
  print(math.max(math.pow(2, 5), 40));
}"#,
        ["40"]
    };

    math_sin_negative_zero => {
        r#"void main() {
  print(math.sin(-0.0));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['-0.0'].
        ["-0.0"]
    };

    math_cos_negative_zero => {
        r#"void main() {
  print(math.cos(-0.0));
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['1.0'].
        ["1.0"]
    };

    math_sqrt_of_fraction => {
        r#"void main() {
  print(math.sqrt(0.25));
}"#,
        ["0.5"]
    };

    math_log_of_pow_returns_exponent => {
        r#"void main() {
  print(math.log(math.pow(2, 8)) > 5.99);
}"#,
        // Corrected against the Dart SDK: this program prints
        // ['false'].
        ["false"]
    };

    math_atan2_on_negative_axis => {
        r#"void main() {
  print(math.atan2(-1, 0) < -1.5);
}"#,
        ["true"]
    };

    math_exp_of_log_roundtrip => {
        r#"void main() {
  print(math.exp(math.log(7)) > 6.99);
}"#,
        ["true"]
    };
}
