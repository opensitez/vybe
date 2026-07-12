//! BigInt: construction, parse, arithmetic, compareTo, toString, parity, gcd.

dart_cases! {
    bigint_from_zero => {
        r#"void main() {
  print(BigInt.from(0));
}"#,
        ["0"]
    };

    bigint_from_positive_int => {
        r#"void main() {
  print(BigInt.from(42));
}"#,
        ["42"]
    };

    bigint_from_negative_int => {
        r#"void main() {
  print(BigInt.from(-17));
}"#,
        ["-17"]
    };

    bigint_from_large_int => {
        r#"void main() {
  print(BigInt.from(999999999999));
}"#,
        ["999999999999"]
    };

    bigint_from_double_whole => {
        r#"void main() {
  print(BigInt.from(100.0));
}"#,
        ["100"]
    };

    bigint_from_double_truncates_fraction => {
        r#"void main() {
  print(BigInt.from(9.9));
}"#,
        ["9"]
    };

    bigint_parse_decimal_string => {
        r#"void main() {
  print(BigInt.parse('12345'));
}"#,
        ["12345"]
    };

    bigint_parse_negative_decimal => {
        r#"void main() {
  print(BigInt.parse('-9876'));
}"#,
        ["-9876"]
    };

    bigint_parse_hex_radix => {
        r#"void main() {
  print(BigInt.parse('FF', radix: 16));
}"#,
        ["255"]
    };

    bigint_parse_binary_radix => {
        r#"void main() {
  print(BigInt.parse('1010', radix: 2));
}"#,
        ["10"]
    };

    bigint_parse_octal_radix => {
        r#"void main() {
  print(BigInt.parse('777', radix: 8));
}"#,
        ["511"]
    };

    bigint_addition_two_positives => {
        r#"void main() {
  var a = BigInt.from(100);
  var b = BigInt.from(23);
  print(a + b);
}"#,
        ["123"]
    };

    bigint_addition_with_negative => {
        r#"void main() {
  var a = BigInt.from(50);
  var b = BigInt.from(-20);
  print(a + b);
}"#,
        ["30"]
    };

    bigint_subtraction_basic => {
        r#"void main() {
  var a = BigInt.from(100);
  var b = BigInt.from(37);
  print(a - b);
}"#,
        ["63"]
    };

    bigint_subtraction_yields_negative => {
        r#"void main() {
  var a = BigInt.from(5);
  var b = BigInt.from(12);
  print(a - b);
}"#,
        ["-7"]
    };

    bigint_multiplication_basic => {
        r#"void main() {
  var a = BigInt.from(12);
  var b = BigInt.from(11);
  print(a * b);
}"#,
        ["132"]
    };

    bigint_multiplication_by_zero => {
        r#"void main() {
  var a = BigInt.from(999);
  print(a * BigInt.from(0));
}"#,
        ["0"]
    };

    bigint_multiplication_negative_operands => {
        r#"void main() {
  var a = BigInt.from(-6);
  var b = BigInt.from(7);
  print(a * b);
}"#,
        ["-42"]
    };

    bigint_truncating_division_whole => {
        r#"void main() {
  var a = BigInt.from(100);
  var b = BigInt.from(4);
  print(a ~/ b);
}"#,
        ["25"]
    };

    bigint_truncating_division_truncates => {
        r#"void main() {
  var a = BigInt.from(17);
  var b = BigInt.from(5);
  print(a ~/ b);
}"#,
        ["3"]
    };

    bigint_truncating_division_negative => {
        r#"void main() {
  var a = BigInt.from(-17);
  var b = BigInt.from(5);
  print(a ~/ b);
}"#,
        ["-3"]
    };

    bigint_modulo_positive => {
        r#"void main() {
  var a = BigInt.from(17);
  var b = BigInt.from(5);
  print(a % b);
}"#,
        ["2"]
    };

    bigint_modulo_negative_dividend => {
        r#"void main() {
  var a = BigInt.from(-17);
  var b = BigInt.from(5);
  print(a % b);
}"#,
        ["3"]
    };

    bigint_modulo_by_one => {
        r#"void main() {
  var a = BigInt.from(42);
  print(a % BigInt.one);
}"#,
        ["0"]
    };

    bigint_compare_to_equal => {
        r#"void main() {
  var a = BigInt.from(10);
  var b = BigInt.from(10);
  print(a.compareTo(b));
}"#,
        ["0"]
    };

    bigint_compare_to_less_than => {
        r#"void main() {
  var a = BigInt.from(3);
  var b = BigInt.from(7);
  print(a.compareTo(b));
}"#,
        ["-1"]
    };

    bigint_compare_to_greater_than => {
        r#"void main() {
  var a = BigInt.from(20);
  var b = BigInt.from(5);
  print(a.compareTo(b));
}"#,
        ["1"]
    };

    bigint_compare_to_negative_vs_positive => {
        r#"void main() {
  var a = BigInt.from(-1);
  var b = BigInt.from(1);
  print(a.compareTo(b));
}"#,
        ["-1"]
    };

    bigint_to_string_decimal => {
        r#"void main() {
  print(BigInt.from(1234567890).toString());
}"#,
        ["1234567890"]
    };

    bigint_to_string_negative => {
        r#"void main() {
  print(BigInt.from(-555).toString());
}"#,
        ["-555"]
    };

    bigint_to_string_zero => {
        r#"void main() {
  print(BigInt.zero.toString());
}"#,
        ["0"]
    };

    bigint_is_even_true => {
        r#"void main() {
  print(BigInt.from(24).isEven);
}"#,
        ["true"]
    };

    bigint_is_even_false => {
        r#"void main() {
  print(BigInt.from(25).isEven);
}"#,
        ["false"]
    };

    bigint_is_odd_true => {
        r#"void main() {
  print(BigInt.from(33).isOdd);
}"#,
        ["true"]
    };

    bigint_is_odd_false => {
        r#"void main() {
  print(BigInt.from(44).isOdd);
}"#,
        ["false"]
    };

    bigint_is_even_zero => {
        r#"void main() {
  print(BigInt.zero.isEven);
}"#,
        ["true"]
    };

    bigint_is_odd_zero => {
        r#"void main() {
  print(BigInt.zero.isOdd);
}"#,
        ["false"]
    };

    bigint_gcd_coprime => {
        r#"void main() {
  var a = BigInt.from(17);
  var b = BigInt.from(13);
  print(BigInt.gcd(a, b));
}"#,
        ["1"]
    };

    bigint_gcd_common_factor => {
        r#"void main() {
  var a = BigInt.from(48);
  var b = BigInt.from(18);
  print(BigInt.gcd(a, b));
}"#,
        ["6"]
    };

    bigint_gcd_one_operand_zero => {
        r#"void main() {
  var a = BigInt.from(0);
  var b = BigInt.from(15);
  print(BigInt.gcd(a, b));
}"#,
        ["15"]
    };

    bigint_gcd_both_zero => {
        r#"void main() {
  print(BigInt.gcd(BigInt.zero, BigInt.zero));
}"#,
        ["0"]
    };

    bigint_gcd_negative_operands => {
        r#"void main() {
  var a = BigInt.from(-36);
  var b = BigInt.from(24);
  print(BigInt.gcd(a, b));
}"#,
        ["12"]
    };

    bigint_unary_minus => {
        r#"void main() {
  var a = BigInt.from(42);
  print(-a);
}"#,
        ["-42"]
    };

    bigint_unary_minus_on_negative => {
        r#"void main() {
  var a = BigInt.from(-7);
  print(-a);
}"#,
        ["7"]
    };

    bigint_abs_positive => {
        r#"void main() {
  print(BigInt.from(99).abs());
}"#,
        ["99"]
    };

    bigint_abs_negative => {
        r#"void main() {
  print(BigInt.from(-99).abs());
}"#,
        ["99"]
    };

    bigint_power_small => {
        r#"void main() {
  var base = BigInt.from(2);
  print(base * base * base);
}"#,
        ["8"]
    };

    bigint_large_addition_no_overflow => {
        r#"void main() {
  var a = BigInt.parse('9007199254740992');
  var b = BigInt.from(1);
  print(a + b);
}"#,
        ["9007199254740993"]
    };

    bigint_equality_same_value => {
        r#"void main() {
  var a = BigInt.from(100);
  var b = BigInt.from(100);
  print(a == b);
}"#,
        ["true"]
    };

    bigint_inequality_different_value => {
        r#"void main() {
  var a = BigInt.from(100);
  var b = BigInt.from(101);
  print(a != b);
}"#,
        ["true"]
    };
}
