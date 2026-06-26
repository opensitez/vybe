//! Dart logical operators: &&, ||, !, short-circuit via side-effect prints,
//! and De Morgan equivalence patterns.

dart_cases! {
    logical_and_both_operands_true => {
        r#"void main() {
  print(true && true);
}"#,
        ["true"]
    };

    logical_and_left_operand_false => {
        r#"void main() {
  print(false && true);
}"#,
        ["false"]
    };

    logical_and_right_operand_false => {
        r#"void main() {
  print(true && false);
}"#,
        ["false"]
    };

    logical_or_both_operands_false => {
        r#"void main() {
  print(false || false);
}"#,
        ["false"]
    };

    logical_or_left_operand_true => {
        r#"void main() {
  print(true || false);
}"#,
        ["true"]
    };

    logical_or_right_operand_true => {
        r#"void main() {
  print(false || true);
}"#,
        ["true"]
    };

    logical_not_on_true => {
        r#"void main() {
  print(!true);
}"#,
        ["false"]
    };

    logical_not_on_false => {
        r#"void main() {
  print(!false);
}"#,
        ["true"]
    };

    double_logical_not_restores_boolean => {
        r#"void main() {
  print(!!true);
}"#,
        ["true"]
    };

    short_circuit_and_skips_rhs_increment => {
        r#"void main() {
  var steps = 0;
  false && (steps = steps + 1) == 1;
  print(steps);
}"#,
        ["0"]
    };

    short_circuit_or_skips_rhs_increment => {
        r#"void main() {
  var steps = 0;
  true || (steps = steps + 1) == 1;
  print(steps);
}"#,
        ["0"]
    };

    short_circuit_and_evaluates_rhs_when_lhs_true => {
        r#"void main() {
  var steps = 0;
  true && (steps = steps + 1) == 1;
  print(steps);
}"#,
        ["1"]
    };

    short_circuit_or_evaluates_rhs_when_lhs_false => {
        r#"void main() {
  var steps = 0;
  false || (steps = steps + 1) == 1;
  print(steps);
}"#,
        ["1"]
    };

    de_morgan_not_of_and_matches_or_of_nots => {
        r#"void main() {
  var a = true;
  var b = false;
  print(!(a && b));
  print(!a || !b);
}"#,
        ["true", "true"]
    };

    de_morgan_not_of_or_matches_and_of_nots => {
        r#"void main() {
  var a = true;
  var b = false;
  print(!(a || b));
  print(!a && !b);
}"#,
        ["false", "false"]
    };

    de_morgan_not_and_with_both_false => {
        r#"void main() {
  var a = false;
  var b = false;
  print(!(a && b));
  print(!a || !b);
}"#,
        ["true", "true"]
    };

    de_morgan_not_or_with_both_true => {
        r#"void main() {
  var a = true;
  var b = true;
  print(!(a || b));
  print(!a && !b);
}"#,
        ["false", "false"]
    };

    and_has_higher_precedence_than_or => {
        r#"void main() {
  print(false || true && false);
}"#,
        ["false"]
    };

    not_binds_tighter_than_and => {
        r#"void main() {
  print(!false && true);
}"#,
        ["true"]
    };

    mixed_and_or_with_parentheses => {
        r#"void main() {
  print((true || false) && false);
}"#,
        ["false"]
    };

    logical_and_with_comparison_operands => {
        r#"void main() {
  print(3 < 5 && 10 > 7);
}"#,
        ["true"]
    };

    logical_or_with_comparison_operands => {
        r#"void main() {
  print(3 > 5 || 10 > 7);
}"#,
        ["true"]
    };

    triple_and_all_true => {
        r#"void main() {
  print(true && true && true);
}"#,
        ["true"]
    };

    triple_and_middle_false_short_circuits => {
        r#"void main() {
  var hit = 0;
  true && false && (hit = hit + 1) == 1;
  print(hit);
}"#,
        ["0"]
    };

    triple_or_all_false => {
        r#"void main() {
  print(false || false || false);
}"#,
        ["false"]
    };

    triple_or_middle_true_short_circuits => {
        r#"void main() {
  var hit = 0;
  false || true || (hit = hit + 1) == 1;
  print(hit);
}"#,
        ["0"]
    };

    not_of_equality_comparison => {
        r#"void main() {
  print(!(5 == 5));
}"#,
        ["false"]
    };

    not_of_inequality_comparison => {
        r#"void main() {
  print(!(5 != 3));
}"#,
        ["false"]
    };

    short_circuit_and_skips_print_side_effect => {
        r#"void main() {
  bool mark() {
    print('rhs');
    return true;
  }
  false && mark();
  print('done');
}"#,
        ["done"]
    };

    short_circuit_or_skips_print_side_effect => {
        r#"void main() {
  bool mark() {
    print('rhs');
    return false;
  }
  true || mark();
  print('done');
}"#,
        ["done"]
    };

    logical_and_result_in_arithmetic_context => {
        r#"void main() {
  print((true && false) ? 1 : 2);
}"#,
        ["2"]
    };

    logical_or_result_in_arithmetic_context => {
        r#"void main() {
  print((false || true) ? 9 : 4);
}"#,
        ["9"]
    };

    chained_not_and_or_expression => {
        r#"void main() {
  print(!false && !false || false);
}"#,
        ["true"]
    };

    logical_and_with_literal_false_operand => {
        r#"void main() {
  print(false && true);
}"#,
        ["false"]
    };

    logical_or_with_literal_true_operand => {
        r#"void main() {
  print(true || false);
}"#,
        ["true"]
    };

    not_of_and_with_comparison_operands => {
        r#"void main() {
  print(!(2 < 1 && 3 < 4));
}"#,
        ["true"]
    };

    not_of_or_with_comparison_operands => {
        r#"void main() {
  print(!(2 < 1 || 3 < 4));
}"#,
        ["false"]
    };

    short_circuit_and_prints_only_lhs_side_effect => {
        r#"void main() {
  bool mark(String label) {
    print(label);
    return true;
  }
  false && mark('rhs');
  print('end');
}"#,
        ["end"]
    };

    short_circuit_or_prints_only_lhs_side_effect => {
        r#"void main() {
  bool mark(String label) {
    print(label);
    return false;
  }
  true || mark('rhs');
  print('end');
}"#,
        ["end"]
    };
}
