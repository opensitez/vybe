//! Dart comparison operators: ==, !=, <, >, <=, >=, strings, doubles,
//! identical references via the same variable, and type promotion.

dart_cases! {
    int_equality_same_values => {
        r#"void main() {
  print(17 == 17);
}"#,
        ["true"]
    };

    int_inequality_different_values => {
        r#"void main() {
  print(17 != 18);
}"#,
        ["true"]
    };

    double_equality_same_fraction => {
        r#"void main() {
  print(3.14 == 3.14);
}"#,
        ["true"]
    };

    int_double_equality_cross_type => {
        r#"void main() {
  print(3 == 3.0);
}"#,
        ["true"]
    };

    int_double_inequality_cross_type => {
        r#"void main() {
  print(3 != 3.1);
}"#,
        ["true"]
    };

    string_equality_same_content => {
        r#"void main() {
  print('dart' == 'dart');
}"#,
        ["true"]
    };

    string_inequality_different_content => {
        r#"void main() {
  print('dart' != 'dart2');
}"#,
        ["true"]
    };

    bool_equality_true_literals => {
        r#"void main() {
  print(true == true);
}"#,
        ["true"]
    };

    bool_equality_false_literals => {
        r#"void main() {
  print(false == false);
}"#,
        ["true"]
    };

    null_equality_null => {
        r#"void main() {
  print(null == null);
}"#,
        ["true"]
    };

    less_than_integers => {
        r#"void main() {
  print(2 < 5);
}"#,
        ["true"]
    };

    greater_than_integers => {
        r#"void main() {
  print(9 > 4);
}"#,
        ["true"]
    };

    less_than_or_equal_equal_values => {
        r#"void main() {
  print(6 <= 6);
}"#,
        ["true"]
    };

    less_than_or_equal_less_values => {
        r#"void main() {
  print(3 <= 7);
}"#,
        ["true"]
    };

    greater_than_or_equal_equal_values => {
        r#"void main() {
  print(11 >= 11);
}"#,
        ["true"]
    };

    greater_than_or_equal_greater_values => {
        r#"void main() {
  print(12 >= 5);
}"#,
        ["true"]
    };

    double_less_than_fractional_values => {
        r#"void main() {
  print(1.5 < 2.5);
}"#,
        ["true"]
    };

    double_greater_than_fractional_values => {
        r#"void main() {
  print(4.2 > 4.1);
}"#,
        ["true"]
    };

    negative_number_less_than_positive => {
        r#"void main() {
  print(-1 < 0);
}"#,
        ["true"]
    };

    string_lexicographic_less => {
        r#"void main() {
  print('abc' < 'abd');
}"#,
        ["true"]
    };

    string_lexicographic_greater => {
        r#"void main() {
  print('zebra' > 'apple');
}"#,
        ["true"]
    };

    string_equality_empty_literals => {
        r#"void main() {
  print('' == '');
}"#,
        ["true"]
    };

    identical_list_via_same_variable => {
        r#"void main() {
  var nums = [1, 2, 3];
  print(nums == nums);
}"#,
        ["true"]
    };

    identical_map_via_same_variable => {
        r#"void main() {
  var scores = {'a': 1};
  print(scores == scores);
}"#,
        ["true"]
    };

    different_list_instances_same_content_equal => {
        r#"void main() {
  print([1, 2] == [1, 2]);
}"#,
        ["true"]
    };

    different_string_instances_same_content_equal => {
        r#"void main() {
  var a = 'dart';
  var b = 'dart';
  print(a == b);
}"#,
        ["true"]
    };

    num_variable_compares_int_to_double => {
        r#"void main() {
  num n = 4;
  print(n < 4.5);
}"#,
        ["true"]
    };

    comparing_int_literal_to_double_literal => {
        r#"void main() {
  print(7 < 7.1);
}"#,
        ["true"]
    };

    equality_false_after_truncated_double => {
        r#"void main() {
  print(3 == 3.2);
}"#,
        ["false"]
    };

    comparing_zero_int_and_double_zero => {
        r#"void main() {
  print(0 == 0.0);
}"#,
        ["true"]
    };

    comparing_negative_zero_double_to_int_zero => {
        r#"void main() {
  print(-0.0 == 0);
}"#,
        ["true"]
    };

    chained_less_and_greater_bounds => {
        r#"void main() {
  var n = 5;
  print(n >= 1 && n <= 10);
}"#,
        ["true"]
    };

    comparison_with_arithmetic_left_operand => {
        r#"void main() {
  print(2 + 3 > 4);
}"#,
        ["true"]
    };

    comparison_with_arithmetic_right_operand => {
        r#"void main() {
  print(10 > 3 * 2);
}"#,
        ["true"]
    };

    string_length_comparison_via_property => {
        r#"void main() {
  print('dart'.length > 2);
}"#,
        ["true"]
    };

    type_promoted_int_comparison_after_is_check => {
        r#"void main() {
  Object? value = 8;
  if (value is int) {
    print(value > 5);
  }
}"#,
        ["true"]
    };

    object_identity_same_instance_via_alias => {
        r#"void main() {
  var original = <int>[];
  var alias = original;
  print(original == alias);
}"#,
        ["true"]
    };

    inequality_is_negation_of_equality => {
        r#"void main() {
  print(5 != 6);
  print(5 == 6);
}"#,
        ["true", "false"]
    };

    comparing_identical_int_literals => {
        r#"void main() {
  print(100 == 100);
}"#,
        ["true"]
    };

    double_less_than_or_equal_cross_type => {
        r#"void main() {
  print(2.5 <= 3);
}"#,
        ["true"]
    };

    double_greater_than_or_equal_cross_type => {
        r#"void main() {
  print(9.0 >= 8);
}"#,
        ["true"]
    };

    string_not_equal_empty_vs_nonempty => {
        r#"void main() {
  print('a' != '');
}"#,
        ["true"]
    };

    list_not_equal_different_lengths => {
        r#"void main() {
  print([1] != [1, 2]);
}"#,
        ["true"]
    };

    comparing_bool_literals_with_not_equal => {
        r#"void main() {
  print(true != false);
}"#,
        ["true"]
    };
}
