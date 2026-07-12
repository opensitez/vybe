//! Dart num/int/double arithmetic: operators, precedence, mixed types,
//! rounding intrinsics, parse, sign, and negative zero.

dart_cases! {
    integer_addition_two_positives => {
        r#"void main() {
  print(2 + 3);
}"#,
        ["5"]
    };

    integer_subtraction_basic => {
        r#"void main() {
  print(10 - 4);
}"#,
        ["6"]
    };

    integer_multiplication_basic => {
        r#"void main() {
  print(6 * 7);
}"#,
        ["42"]
    };

    integer_division_always_produces_double => {
        r#"void main() {
  print(7 / 2);
}"#,
        ["3.5"]
    };

    truncating_integer_division_truncates_toward_zero => {
        r#"void main() {
  print(7 ~/ 2);
}"#,
        ["3"]
    };

    truncating_division_negative_dividend_toward_zero => {
        r#"void main() {
  print(-7 ~/ 2);
}"#,
        ["-3"]
    };

    truncating_division_negative_divisor_toward_zero => {
        r#"void main() {
  print(7 ~/ -2);
}"#,
        ["-3"]
    };

    modulo_positive_operands => {
        r#"void main() {
  print(10 % 3);
}"#,
        ["1"]
    };

    modulo_negative_dividend_matches_dividend_sign => {
        r#"void main() {
  print(-7 % 3);
}"#,
        ["-1"]
    };

    modulo_negative_divisor => {
        r#"void main() {
  print(7 % -3);
}"#,
        ["1"]
    };

    modulo_both_operands_negative => {
        r#"void main() {
  print(-7 % -3);
}"#,
        ["-1"]
    };

    modulo_on_double_operands => {
        r#"void main() {
  print(5.5 % 2);
}"#,
        ["1.5"]
    };

    unary_minus_on_integer_literal => {
        r#"void main() {
  print(-(-8));
}"#,
        ["8"]
    };

    unary_minus_on_double_literal => {
        r#"void main() {
  print(-(-2.5));
}"#,
        ["2.5"]
    };

    multiplication_has_higher_precedence_than_addition => {
        r#"void main() {
  print(2 + 3 * 4);
}"#,
        ["14"]
    };

    multiplication_has_higher_precedence_than_subtraction => {
        r#"void main() {
  print(20 - 3 * 2);
}"#,
        ["14"]
    };

    division_and_modulo_bind_left_to_right => {
        r#"void main() {
  print(20 / 4 / 2);
}"#,
        ["2.5"]
    };

    subtraction_is_left_associative => {
        r#"void main() {
  print(10 - 3 - 2);
}"#,
        ["5"]
    };

    parenthesized_sum_before_multiplication => {
        r#"void main() {
  print((2 + 3) * 4);
}"#,
        ["20"]
    };

    nested_parentheses_control_evaluation_order => {
        r#"void main() {
  print(((1 + 2) * 3) + 4);
}"#,
        ["13"]
    };

    int_plus_double_promotes_to_double => {
        r#"void main() {
  print(2 + 3.0);
}"#,
        ["5.0"]
    };

    int_minus_double_promotes_to_double => {
        r#"void main() {
  print(5 - 2.5);
}"#,
        ["2.5"]
    };

    int_times_double_promotes_to_double => {
        r#"void main() {
  print(2 * 2.5);
}"#,
        ["5.0"]
    };

    int_divided_by_int_yields_double_even_when_whole => {
        r#"void main() {
  print(4 / 2);
}"#,
        ["2.0"]
    };

    double_divided_by_int => {
        r#"void main() {
  print(5.0 / 2);
}"#,
        ["2.5"]
    };

    truncating_division_on_double_operands => {
        r#"void main() {
  print(9.0 ~/ 2.0);
}"#,
        ["4"]
    };

    negative_zero_literal_prints_as_zero => {
        r#"void main() {
  print(-0.0);
}"#,
        ["0.0"]
    };

    negative_zero_plus_positive_becomes_positive => {
        r#"void main() {
  print(-0.0 + 1.0);
}"#,
        ["1.0"]
    };

    abs_on_negative_integer => {
        r#"void main() {
  print((-7).abs());
}"#,
        ["7"]
    };

    abs_on_positive_integer => {
        r#"void main() {
  print(4.abs());
}"#,
        ["4"]
    };

    abs_on_negative_double => {
        r#"void main() {
  print((-2.5).abs());
}"#,
        ["2.5"]
    };

    abs_on_negative_zero_yields_positive_zero => {
        r#"void main() {
  print((-0.0).abs());
}"#,
        ["0.0"]
    };

    round_half_up_on_positive_fraction => {
        r#"void main() {
  print(2.5.round());
}"#,
        ["3"]
    };

    round_down_on_fraction_below_half => {
        r#"void main() {
  print(2.4.round());
}"#,
        ["2"]
    };

    round_on_negative_fraction => {
        r#"void main() {
  print((-2.6).round());
}"#,
        ["-3"]
    };

    floor_truncates_toward_negative_infinity => {
        r#"void main() {
  print(3.9.floor());
}"#,
        ["3"]
    };

    floor_on_negative_fraction => {
        r#"void main() {
  print((-1.2).floor());
}"#,
        ["-2"]
    };

    ceil_rounds_toward_positive_infinity => {
        r#"void main() {
  print(3.1.ceil());
}"#,
        ["4"]
    };

    ceil_on_negative_fraction => {
        r#"void main() {
  print((-1.8).ceil());
}"#,
        ["-1"]
    };

    truncate_drops_fractional_part_toward_zero => {
        r#"void main() {
  print(3.9.truncate());
}"#,
        ["3"]
    };

    truncate_on_negative_fraction => {
        r#"void main() {
  print((-3.9).truncate());
}"#,
        ["-3"]
    };

    remainder_on_integer_operands => {
        r#"void main() {
  print(7.remainder(3));
}"#,
        ["1"]
    };

    remainder_on_double_operands => {
        r#"void main() {
  print(7.5.remainder(2.0));
}"#,
        ["1.5"]
    };

    to_int_truncates_toward_zero => {
        r#"void main() {
  print(3.7.toInt());
}"#,
        ["3"]
    };

    to_int_on_negative_double => {
        r#"void main() {
  print((-3.7).toInt());
}"#,
        ["-3"]
    };

    to_double_from_integer => {
        r#"void main() {
  print(5.toDouble());
}"#,
        ["5.0"]
    };

    int_parse_decimal_string => {
        r#"void main() {
  print(int.parse('42'));
}"#,
        ["42"]
    };

    double_parse_fractional_string => {
        r#"void main() {
  print(double.parse('3.5'));
}"#,
        ["3.5"]
    };

    num_parse_integer_string => {
        r#"void main() {
  print(num.parse('17'));
}"#,
        ["17"]
    };

    is_negative_false_for_positive_number => {
        r#"void main() {
  print(5.isNegative);
}"#,
        ["false"]
    };

    is_negative_true_for_negative_number => {
        r#"void main() {
  print((-3).isNegative);
}"#,
        ["true"]
    };

    is_negative_false_for_negative_zero => {
        r#"void main() {
  print((-0.0).isNegative);
}"#,
        ["false"]
    };

    sign_of_positive_number => {
        r#"void main() {
  print(5.sign);
}"#,
        ["1.0"]
    };

    sign_of_negative_number => {
        r#"void main() {
  print((-4).sign);
}"#,
        ["-1.0"]
    };

    sign_of_zero => {
        r#"void main() {
  print(0.sign);
}"#,
        ["0.0"]
    };

    sign_of_negative_zero => {
        r#"void main() {
  print((-0.0).sign);
}"#,
        ["-0.0"]
    };
}
