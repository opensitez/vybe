//! Conditional (? :) expressions: basic, nested, null-coalesce combos, and in print.

dart_cases! {
    ternary_true_branch_selected => {
        r#"void main() {
  var label = true ? 'yes' : 'no';
  print(label);
}"#,
        ["yes"]
    };

    ternary_false_branch_selected => {
        r#"void main() {
  var label = false ? 'yes' : 'no';
  print(label);
}"#,
        ["no"]
    };

    ternary_with_numeric_comparison => {
        r#"void main() {
  var n = 10;
  print(n > 5 ? 'big' : 'small');
}"#,
        ["big"]
    };

    ternary_with_numeric_comparison_false => {
        r#"void main() {
  var n = 2;
  print(n > 5 ? 'big' : 'small');
}"#,
        ["small"]
    };

    ternary_nested_three_way_grade => {
        r#"void main() {
  var score = 85;
  var grade = score >= 90 ? 'A' : score >= 80 ? 'B' : 'C';
  print(grade);
}"#,
        ["B"]
    };

    ternary_nested_lowest_branch => {
        r#"void main() {
  var score = 60;
  var grade = score >= 90 ? 'A' : score >= 80 ? 'B' : 'C';
  print(grade);
}"#,
        ["C"]
    };

    ternary_nested_highest_branch => {
        r#"void main() {
  var score = 95;
  var grade = score >= 90 ? 'A' : score >= 80 ? 'B' : 'C';
  print(grade);
}"#,
        ["A"]
    };

    ternary_with_null_coalesce_left_null => {
        r#"void main() {
  String? name;
  print(name ?? 'anonymous');
}"#,
        ["anonymous"]
    };

    ternary_with_null_coalesce_left_present => {
        r#"void main() {
  String? name = 'Ada';
  print(name ?? 'anonymous');
}"#,
        ["Ada"]
    };

    ternary_null_check_with_fallback => {
        r#"void main() {
  String? value;
  print(value != null ? value : 'none');
}"#,
        ["none"]
    };

    ternary_null_check_with_value => {
        r#"void main() {
  String? value = 'data';
  print(value != null ? value : 'none');
}"#,
        ["data"]
    };

    ternary_combined_with_null_coalesce => {
        r#"void main() {
  String? a;
  String? b = 'backup';
  print(a ?? b ?? 'empty');
}"#,
        ["backup"]
    };

    ternary_double_null_coalesce_fallback => {
        r#"void main() {
  String? a;
  String? b;
  print(a ?? b ?? 'empty');
}"#,
        ["empty"]
    };

    ternary_in_print_directly => {
        r#"void main() {
  var x = 4;
  print(x % 2 == 0 ? 'even' : 'odd');
}"#,
        ["even"]
    };

    ternary_in_print_odd_result => {
        r#"void main() {
  var x = 5;
  print(x % 2 == 0 ? 'even' : 'odd');
}"#,
        ["odd"]
    };

    ternary_with_string_interpolation => {
        r#"void main() {
  var n = 3;
  print('count: ${n == 1 ? 'one' : 'many'}');
}"#,
        ["count: many"]
    };

    ternary_string_interpolation_singular => {
        r#"void main() {
  var n = 1;
  print('count: ${n == 1 ? 'one' : 'many'}');
}"#,
        ["count: one"]
    };

    ternary_equality_in_condition => {
        r#"void main() {
  var a = 7;
  var b = 7;
  print(a == b ? 'match' : 'diff');
}"#,
        ["match"]
    };

    ternary_inequality_in_condition => {
        r#"void main() {
  var a = 3;
  var b = 8;
  print(a == b ? 'match' : 'diff');
}"#,
        ["diff"]
    };

    ternary_with_logical_and_condition => {
        r#"void main() {
  var age = 25;
  var licensed = true;
  print(age >= 18 && licensed ? 'drive' : 'no drive');
}"#,
        ["drive"]
    };

    ternary_with_logical_or_condition => {
        r#"void main() {
  var rainy = false;
  var snowy = true;
  print(rainy || snowy ? 'wet' : 'dry');
}"#,
        ["wet"]
    };

    ternary_nested_with_null_coalesce => {
        r#"void main() {
  String? primary;
  String? secondary = 'second';
  print(primary != null ? primary : secondary ?? 'missing');
}"#,
        ["second"]
    };

    ternary_nested_null_coalesce_all_null => {
        r#"void main() {
  String? primary;
  String? secondary;
  print(primary != null ? primary : secondary ?? 'missing');
}"#,
        ["missing"]
    };

    ternary_assign_to_variable => {
        r#"void main() {
  var flag = false;
  var result = flag ? 100 : 200;
  print(result);
}"#,
        ["200"]
    };

    ternary_assign_true_branch_number => {
        r#"void main() {
  var flag = true;
  var result = flag ? 100 : 200;
  print(result);
}"#,
        ["100"]
    };

    ternary_in_return_statement => {
        r#"int absVal(int n) {
  return n >= 0 ? n : -n;
}
void main() {
  print(absVal(-5));
}"#,
        ["5"]
    };

    ternary_in_return_positive => {
        r#"int absVal(int n) {
  return n >= 0 ? n : -n;
}
void main() {
  print(absVal(5));
}"#,
        ["5"]
    };

    ternary_chained_three_levels => {
        r#"void main() {
  var n = 15;
  var size = n > 20 ? 'xl' : n > 10 ? 'lg' : n > 5 ? 'md' : 'sm';
  print(size);
}"#,
        ["lg"]
    };

    ternary_chained_smallest_size => {
        r#"void main() {
  var n = 3;
  var size = n > 20 ? 'xl' : n > 10 ? 'lg' : n > 5 ? 'md' : 'sm';
  print(size);
}"#,
        ["sm"]
    };

    ternary_with_modulo_condition => {
        r#"void main() {
  var n = 9;
  print(n % 3 == 0 ? 'div3' : 'not3');
}"#,
        ["div3"]
    };

    ternary_with_negative_number => {
        r#"void main() {
  var n = -4;
  print(n < 0 ? 'neg' : 'pos');
}"#,
        ["neg"]
    };

    ternary_parenthesized_condition => {
        r#"void main() {
  var a = 2;
  var b = 3;
  print((a + b) > 4 ? 'sum-big' : 'sum-small');
}"#,
        ["sum-big"]
    };
}
