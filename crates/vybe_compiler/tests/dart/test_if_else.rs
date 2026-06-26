//! If statements, else clauses, else-if chains, nested conditionals, and
//! boolean / relational / logical operators in conditions.

dart_cases! {
    if_statement_true_condition_executes_then => {
        r#"void main() {
  if (true) {
    print('then');
  }
}"#,
        ["then"]
    };

    if_statement_false_condition_skips_then => {
        r#"void main() {
  if (false) {
    print('then');
  }
  print('after');
}"#,
        ["after"]
    };

    if_statement_without_else_when_true => {
        r#"void main() {
  var flag = true;
  if (flag) {
    print('yes');
  }
  print('done');
}"#,
        ["yes", "done"]
    };

    if_statement_without_else_when_false => {
        r#"void main() {
  var flag = false;
  if (flag) {
    print('yes');
  }
  print('done');
}"#,
        ["done"]
    };

    else_clause_executes_when_condition_false => {
        r#"void main() {
  if (false) {
    print('then');
  } else {
    print('else');
  }
}"#,
        ["else"]
    };

    else_clause_skipped_when_condition_true => {
        r#"void main() {
  if (true) {
    print('then');
  } else {
    print('else');
  }
}"#,
        ["then"]
    };

    else_if_chain_matches_first_alternative => {
        r#"void main() {
  var n = 95;
  if (n >= 90) {
    print('A');
  } else if (n >= 80) {
    print('B');
  } else if (n >= 70) {
    print('C');
  } else {
    print('F');
  }
}"#,
        ["A"]
    };

    else_if_chain_matches_middle_alternative => {
        r#"void main() {
  var n = 85;
  if (n >= 90) {
    print('A');
  } else if (n >= 80) {
    print('B');
  } else if (n >= 70) {
    print('C');
  } else {
    print('F');
  }
}"#,
        ["B"]
    };

    else_if_chain_matches_last_else_if => {
        r#"void main() {
  var n = 72;
  if (n >= 90) {
    print('A');
  } else if (n >= 80) {
    print('B');
  } else if (n >= 70) {
    print('C');
  } else {
    print('F');
  }
}"#,
        ["C"]
    };

    else_if_chain_falls_through_to_else => {
        r#"void main() {
  var n = 55;
  if (n >= 90) {
    print('A');
  } else if (n >= 80) {
    print('B');
  } else if (n >= 70) {
    print('C');
  } else {
    print('F');
  }
}"#,
        ["F"]
    };

    else_if_chain_without_final_else => {
        r#"void main() {
  var n = 50;
  if (n >= 90) {
    print('A');
  } else if (n >= 80) {
    print('B');
  }
  print('end');
}"#,
        ["end"]
    };

    nested_if_both_conditions_true => {
        r#"void main() {
  var x = 3;
  var y = 7;
  if (x > 0) {
    if (y > 0) {
      print('both positive');
    }
  }
}"#,
        ["both positive"]
    };

    nested_if_outer_true_inner_false => {
        r#"void main() {
  var x = 5;
  var y = -1;
  if (x > 0) {
    if (y > 0) {
      print('inner-then');
    } else {
      print('inner-else');
    }
  }
}"#,
        ["inner-else"]
    };

    nested_if_outer_false_skips_inner => {
        r#"void main() {
  var x = -1;
  var y = 10;
  if (x > 0) {
    if (y > 0) {
      print('inner');
    }
  } else {
    print('outer-else');
  }
}"#,
        ["outer-else"]
    };

    nested_if_three_levels_deep => {
        r#"void main() {
  var a = 1;
  var b = 2;
  var c = 3;
  if (a == 1) {
    if (b == 2) {
      if (c == 3) {
        print('deep');
      }
    }
  }
}"#,
        ["deep"]
    };

    dangling_else_binds_to_nearest_if => {
        r#"void main() {
  if (true) {
    if (false) {
      print('inner-then');
    } else {
      print('inner-else');
    }
  }
}"#,
        ["inner-else"]
    };

    dangling_else_if_associates_with_outer_if => {
        r#"void main() {
  if (false)
    print('skip');
  else if (true)
    print('else-if');
  else
    print('else');
  print('done');
}"#,
        ["else-if", "done"]
    };

    consecutive_independent_if_statements => {
        r#"void main() {
  if (true) {
    print('first');
  }
  if (false) {
    print('skip');
  }
  if (true) {
    print('second');
  }
}"#,
        ["first", "second"]
    };

    boolean_literal_true_in_condition => {
        r#"void main() {
  if (true) {
    print('lit-true');
  }
}"#,
        ["lit-true"]
    };

    boolean_literal_false_in_condition => {
        r#"void main() {
  if (false) {
    print('lit-false');
  } else {
    print('not-false');
  }
}"#,
        ["not-false"]
    };

    boolean_variable_in_condition => {
        r#"void main() {
  var ready = true;
  if (ready) {
    print('ready');
  } else {
    print('not-ready');
  }
}"#,
        ["ready"]
    };

    negated_boolean_in_condition => {
        r#"void main() {
  var closed = false;
  if (!closed) {
    print('open');
  } else {
    print('closed');
  }
}"#,
        ["open"]
    };

    equality_operator_in_condition => {
        r#"void main() {
  var x = 42;
  if (x == 42) {
    print('equal');
  } else {
    print('not-equal');
  }
}"#,
        ["equal"]
    };

    inequality_operator_in_condition => {
        r#"void main() {
  var x = 7;
  if (x != 3) {
    print('different');
  } else {
    print('same');
  }
}"#,
        ["different"]
    };

    less_than_relational_in_condition => {
        r#"void main() {
  var x = 2;
  if (x < 5) {
    print('below');
  } else {
    print('above');
  }
}"#,
        ["below"]
    };

    greater_than_or_equal_relational_in_condition => {
        r#"void main() {
  var x = 10;
  if (x >= 10) {
    print('at-least');
  } else {
    print('below');
  }
}"#,
        ["at-least"]
    };

    logical_and_both_operands_true => {
        r#"void main() {
  var a = true;
  var b = true;
  if (a && b) {
    print('and-true');
  } else {
    print('and-false');
  }
}"#,
        ["and-true"]
    };

    logical_and_short_circuits_on_false => {
        r#"void main() {
  var a = false;
  var b = true;
  if (a && b) {
    print('and-true');
  } else {
    print('and-false');
  }
}"#,
        ["and-false"]
    };

    logical_or_short_circuits_on_true => {
        r#"void main() {
  var a = true;
  var b = false;
  if (a || b) {
    print('or-true');
  } else {
    print('or-false');
  }
}"#,
        ["or-true"]
    };

    logical_or_both_operands_false => {
        r#"void main() {
  var a = false;
  var b = false;
  if (a || b) {
    print('or-true');
  } else {
    print('or-false');
  }
}"#,
        ["or-false"]
    };

    logical_not_operator_in_condition => {
        r#"void main() {
  var active = false;
  if (!(active)) {
    print('inactive');
  } else {
    print('active');
  }
}"#,
        ["inactive"]
    };

    mixed_logical_operators_with_precedence => {
        r#"void main() {
  var p = true;
  var q = false;
  var r = true;
  if (p || q && r) {
    print('match');
  } else {
    print('no-match');
  }
}"#,
        ["match"]
    };

    if_assigns_variable_in_then_branch => {
        r#"void main() {
  var result = 0;
  if (true) {
    result = 10;
  }
  print(result);
}"#,
        ["10"]
    };

    if_assigns_variable_in_else_branch => {
        r#"void main() {
  var result = 0;
  if (false) {
    result = 10;
  } else {
    result = 20;
  }
  print(result);
}"#,
        ["20"]
    };

    if_modifies_existing_variable_both_branches => {
        r#"void main() {
  var sign = 'zero';
  var n = -3;
  if (n > 0) {
    sign = 'positive';
  } else if (n < 0) {
    sign = 'negative';
  } else {
    sign = 'zero';
  }
  print(sign);
}"#,
        ["negative"]
    };

    string_equality_in_condition => {
        r#"void main() {
  var name = 'dart';
  if (name == 'dart') {
    print('match');
  } else {
    print('no-match');
  }
}"#,
        ["match"]
    };

    string_inequality_in_condition => {
        r#"void main() {
  var name = 'vybe';
  if (name != 'dart') {
    print('other');
  } else {
    print('dart');
  }
}"#,
        ["other"]
    };

    null_equality_check_in_condition => {
        r#"void main() {
  String? value = null;
  if (value == null) {
    print('null');
  } else {
    print('non-null');
  }
}"#,
        ["null"]
    };

    is_type_test_promotes_in_then_branch => {
        r#"void main() {
  Object value = 'hello';
  if (value is String) {
    print(value.length);
  } else {
    print(-1);
  }
}"#,
        ["5"]
    };

    is_not_type_test_in_condition => {
        r#"void main() {
  Object value = 42;
  if (value is! String) {
    print('not-string');
  } else {
    print('string');
  }
}"#,
        ["not-string"]
    };

    modulo_in_arithmetic_condition => {
        r#"void main() {
  var n = 10;
  if (n % 3 == 1) {
    print('remainder-one');
  } else {
    print('other-remainder');
  }
}"#,
        ["remainder-one"]
    };

    parenthesized_complex_condition => {
        r#"void main() {
  var x = 4;
  var y = 6;
  if ((x + y) > 8 && (x * y) < 30) {
    print('complex-true');
  } else {
    print('complex-false');
  }
}"#,
        ["complex-true"]
    };

    if_else_selects_branch_via_function_side_effect => {
        r#"void main() {
  var log = '';
  void record(String msg) { log = msg; }
  if (2 + 2 == 4) {
    record('math-ok');
  } else {
    record('math-fail');
  }
  print(log);
}"#,
        ["math-ok"]
    };
}
