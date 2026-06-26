//! Boolean conditions that assert would verify, printing true/false instead of using assert.

dart_cases! {
    assert_like_equality_true => {
        r#"void main() {
  print(1 + 1 == 2);
}"#,
        ["true"]
    };

    assert_like_equality_false => {
        r#"void main() {
  print(1 + 1 == 3);
}"#,
        ["false"]
    };

    assert_like_greater_than => {
        r#"void main() {
  print(5 > 3);
}"#,
        ["true"]
    };

    assert_like_less_than => {
        r#"void main() {
  print(2 < 10);
}"#,
        ["true"]
    };

    assert_like_less_than_false => {
        r#"void main() {
  print(10 < 2);
}"#,
        ["false"]
    };

    assert_like_not_equal => {
        r#"void main() {
  print(7 != 4);
}"#,
        ["true"]
    };

    assert_like_string_equality => {
        r#"void main() {
  print('dart' == 'dart');
}"#,
        ["true"]
    };

    assert_like_string_inequality => {
        r#"void main() {
  print('foo' == 'bar');
}"#,
        ["false"]
    };

    assert_like_logical_and_both_true => {
        r#"void main() {
  print(true && true);
}"#,
        ["true"]
    };

    assert_like_logical_and_one_false => {
        r#"void main() {
  print(true && false);
}"#,
        ["false"]
    };

    assert_like_logical_or_one_true => {
        r#"void main() {
  print(false || true);
}"#,
        ["true"]
    };

    assert_like_logical_or_both_false => {
        r#"void main() {
  print(false || false);
}"#,
        ["false"]
    };

    assert_like_logical_not => {
        r#"void main() {
  print(!false);
}"#,
        ["true"]
    };

    assert_like_null_is_null => {
        r#"void main() {
  String? s;
  print(s == null);
}"#,
        ["true"]
    };

    assert_like_null_is_not_null => {
        r#"void main() {
  String? s = 'set';
  print(s != null);
}"#,
        ["true"]
    };

    assert_like_list_length => {
        r#"void main() {
  var list = [1, 2, 3];
  print(list.length == 3);
}"#,
        ["true"]
    };

    assert_like_list_contains => {
        r#"void main() {
  var list = [1, 2, 3];
  print(list.contains(2));
}"#,
        ["true"]
    };

    assert_like_list_not_contains => {
        r#"void main() {
  var list = [1, 2, 3];
  print(list.contains(9));
}"#,
        ["false"]
    };

    assert_like_arithmetic_sum => {
        r#"void main() {
  print(10 + 5 == 15);
}"#,
        ["true"]
    };

    assert_like_arithmetic_product => {
        r#"void main() {
  print(6 * 7 == 42);
}"#,
        ["true"]
    };

    assert_like_modulo_zero => {
        r#"void main() {
  print(10 % 5 == 0);
}"#,
        ["true"]
    };

    assert_like_chained_comparison => {
        r#"void main() {
  var n = 5;
  print(n > 0 && n < 10);
}"#,
        ["true"]
    };

    assert_like_string_starts_with => {
        r#"void main() {
  print('hello'.startsWith('he'));
}"#,
        ["true"]
    };

    assert_like_string_ends_with => {
        r#"void main() {
  print('hello'.endsWith('lo'));
}"#,
        ["true"]
    };

    assert_like_is_empty_list => {
        r#"void main() {
  var list = <int>[];
  print(list.isEmpty);
}"#,
        ["true"]
    };

    assert_like_is_not_empty_list => {
        r#"void main() {
  var list = [0];
  print(list.isNotEmpty);
}"#,
        ["true"]
    };

    assert_like_identity_same_list_ref => {
        r#"void main() {
  var a = [1];
  var b = a;
  print(a == b);
}"#,
        ["true"]
    };

    assert_like_type_check_int => {
        r#"void main() {
  var n = 42;
  print(n is int);
}"#,
        ["true"]
    };

    assert_like_type_check_not_string => {
        r#"void main() {
  var n = 42;
  print(n is String);
}"#,
        ["false"]
    };

    assert_like_combined_math_and_logic => {
        r#"void main() {
  print((3 + 4) == 7 && (2 * 3) == 6);
}"#,
        ["true"]
    };

    assert_like_negation_of_false_equality => {
        r#"void main() {
  print(!(5 == 5) == false);
}"#,
        ["true"]
    };
}
