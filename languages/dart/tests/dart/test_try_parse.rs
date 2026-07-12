//! int.tryParse, double.tryParse, int.parse throws, radix parsing, invalid null.

dart_cases! {
    int_try_parse_valid_decimal => {
        r#"void main() {
  print(int.tryParse('42'));
}"#,
        ["42"]
    };

    int_try_parse_negative_decimal => {
        r#"void main() {
  print(int.tryParse('-17'));
}"#,
        ["-17"]
    };

    int_try_parse_zero => {
        r#"void main() {
  print(int.tryParse('0'));
}"#,
        ["0"]
    };

    int_try_parse_leading_zeros => {
        r#"void main() {
  print(int.tryParse('007'));
}"#,
        ["7"]
    };

    int_try_parse_invalid_letters => {
        r#"void main() {
  print(int.tryParse('abc'));
}"#,
        ["null"]
    };

    int_try_parse_empty_string => {
        r#"void main() {
  print(int.tryParse(''));
}"#,
        ["null"]
    };

    int_try_parse_whitespace_only => {
        r#"void main() {
  print(int.tryParse('   '));
}"#,
        ["null"]
    };

    int_try_parse_mixed_alphanumeric => {
        r#"void main() {
  print(int.tryParse('12abc'));
}"#,
        ["null"]
    };

    int_try_parse_hex_ff_radix_16 => {
        r#"void main() {
  print(int.tryParse('FF', radix: 16));
}"#,
        ["255"]
    };

    int_try_parse_hex_lowercase_radix_16 => {
        r#"void main() {
  print(int.tryParse('ff', radix: 16));
}"#,
        ["255"]
    };

    int_try_parse_octal_77_radix_8 => {
        r#"void main() {
  print(int.tryParse('77', radix: 8));
}"#,
        ["63"]
    };

    int_try_parse_binary_1010_radix_2 => {
        r#"void main() {
  print(int.tryParse('1010', radix: 2));
}"#,
        ["10"]
    };

    int_try_parse_hex_with_0x_prefix_invalid => {
        r#"void main() {
  print(int.tryParse('0xFF', radix: 16));
}"#,
        ["null"]
    };

    int_try_parse_decimal_with_dot => {
        r#"void main() {
  print(int.tryParse('3.14'));
}"#,
        ["null"]
    };

    int_try_parse_plus_sign_prefix => {
        r#"void main() {
  print(int.tryParse('+25'));
}"#,
        ["25"]
    };

    int_try_parse_large_valid => {
        r#"void main() {
  print(int.tryParse('2147483647'));
}"#,
        ["2147483647"]
    };

    double_try_parse_valid => {
        r#"void main() {
  print(double.tryParse('3.14'));
}"#,
        ["3.14"]
    };

    double_try_parse_negative => {
        r#"void main() {
  print(double.tryParse('-2.5'));
}"#,
        ["-2.5"]
    };

    double_try_parse_integer_string => {
        r#"void main() {
  print(double.tryParse('100'));
}"#,
        ["100.0"]
    };

    double_try_parse_scientific_notation => {
        r#"void main() {
  print(double.tryParse('1.5e2'));
}"#,
        ["150.0"]
    };

    double_try_parse_invalid => {
        r#"void main() {
  print(double.tryParse('not-a-number'));
}"#,
        ["null"]
    };

    double_try_parse_empty => {
        r#"void main() {
  print(double.tryParse(''));
}"#,
        ["null"]
    };

    double_try_parse_nan_string => {
        r#"void main() {
  print(double.tryParse('NaN'));
}"#,
        ["NaN"]
    };

    double_try_parse_infinity_string => {
        r#"void main() {
  print(double.tryParse('Infinity'));
}"#,
        ["Infinity"]
    };

    double_try_parse_negative_infinity_string => {
        r#"void main() {
  print(double.tryParse('-Infinity'));
}"#,
        ["-Infinity"]
    };

    int_parse_throws_on_invalid_caught => {
        r#"void main() {
  try {
    int.parse('xyz');
  } catch (e) {
    print('caught');
  }
}"#,
        ["caught"]
    };

    int_parse_throws_on_empty_caught => {
        r#"void main() {
  try {
    int.parse('');
  } catch (e) {
    print('caught');
  }
}"#,
        ["caught"]
    };

    int_parse_throws_on_fraction_caught => {
        r#"void main() {
  try {
    int.parse('3.14');
  } catch (e) {
    print('caught');
  }
}"#,
        ["caught"]
    };

    int_parse_valid_does_not_throw => {
        r#"void main() {
  try {
    print(int.parse('99'));
  } catch (e) {
    print('caught');
  }
}"#,
        ["99"]
    };

    int_parse_hex_radix_16 => {
        r#"void main() {
  print(int.parse('A', radix: 16));
}"#,
        ["10"]
    };

    int_parse_octal_radix_8 => {
        r#"void main() {
  print(int.parse('10', radix: 8));
}"#,
        ["8"]
    };

    int_parse_binary_radix_2 => {
        r#"void main() {
  print(int.parse('1111', radix: 2));
}"#,
        ["15"]
    };

    int_parse_negative_with_radix => {
        r#"void main() {
  print(int.parse('-F', radix: 16));
}"#,
        ["-15"]
    };

    int_try_parse_radix_16_invalid_digit => {
        r#"void main() {
  print(int.tryParse('G', radix: 16));
}"#,
        ["null"]
    };

    int_try_parse_radix_2_invalid_digit => {
        r#"void main() {
  print(int.tryParse('2', radix: 2));
}"#,
        ["null"]
    };

    double_try_parse_trailing_garbage => {
        r#"void main() {
  print(double.tryParse('3.14px'));
}"#,
        ["null"]
    };

    int_try_parse_trailing_garbage => {
        r#"void main() {
  print(int.tryParse('42px'));
}"#,
        ["null"]
    };

    try_parse_null_coalesce_default => {
        r#"void main() {
  print(int.tryParse('bad') ?? -1);
}"#,
        ["-1"]
    };

    double_try_parse_null_coalesce => {
        r#"void main() {
  print(double.tryParse('bad') ?? 0.0);
}"#,
        ["0.0"]
    };

    int_try_parse_then_add => {
        r#"void main() {
  var n = int.tryParse('10');
  print(n! + 5);
}"#,
        ["15"]
    };

    double_try_parse_then_multiply => {
        r#"void main() {
  var d = double.tryParse('2.5');
  print(d! * 4);
}"#,
        ["10.0"]
    };

    int_parse_throws_format_exception_type => {
        r#"void main() {
  try {
    int.parse('oops');
  } on FormatException {
    print('format');
  }
}"#,
        ["format"]
    };

    int_try_parse_base36 => {
        r#"void main() {
  print(int.tryParse('Z', radix: 36));
}"#,
        ["35"]
    };

    int_try_parse_base36_ten => {
        r#"void main() {
  print(int.tryParse('10', radix: 36));
}"#,
        ["36"]
    };

    double_try_parse_leading_plus => {
        r#"void main() {
  print(double.tryParse('+3.5'));
}"#,
        ["3.5"]
    };

    int_try_parse_only_minus_sign => {
        r#"void main() {
  print(int.tryParse('-'));
}"#,
        ["null"]
    };

    double_try_parse_only_dot => {
        r#"void main() {
  print(double.tryParse('.'));
}"#,
        ["null"]
    };

    int_parse_radix_16_uppercase => {
        r#"void main() {
  print(int.parse('DEAD', radix: 16));
}"#,
        ["57005"]
    };

    int_try_parse_binary_all_ones => {
        r#"void main() {
  print(int.tryParse('11111111', radix: 2));
}"#,
        ["255"]
    };

    double_parse_throws_on_invalid_caught => {
        r#"void main() {
  try {
    double.parse('not-a-double');
  } catch (e) {
    print('caught');
  }
}"#,
        ["caught"]
    };
}
