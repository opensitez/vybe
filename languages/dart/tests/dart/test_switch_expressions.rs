//! Switch expressions (=> arms): int/String yields, nested expressions,
//! and using switch as a value in assignments, returns, and calls.

dart_cases! {
    switch_expr_yields_int_for_case_one => {
        r#"void main() {
  var code = 1;
  var points = switch (code) {
    1 => 10,
    2 => 20,
    _ => 0,
  };
  print(points);
}"#,
        ["10"]
    };

    switch_expr_yields_int_for_case_two => {
        r#"void main() {
  var code = 2;
  var points = switch (code) {
    1 => 10,
    2 => 20,
    _ => 0,
  };
  print(points);
}"#,
        ["20"]
    };

    switch_expr_yields_int_wildcard_default => {
        r#"void main() {
  var code = 99;
  var points = switch (code) {
    1 => 10,
    2 => 20,
    _ => 0,
  };
  print(points);
}"#,
        ["0"]
    };

    switch_expr_yields_int_zero_literal => {
        r#"void main() {
  var n = switch (0) {
    0 => 100,
    _ => -1,
  };
  print(n);
}"#,
        ["100"]
    };

    switch_expr_yields_int_negative_selector => {
        r#"void main() {
  var n = switch (-2) {
    -2 => 42,
    _ => 0,
  };
  print(n);
}"#,
        ["42"]
    };

    switch_expr_yields_int_sum_in_arm => {
        r#"void main() {
  var tier = 3;
  var bonus = switch (tier) {
    1 => 5 + 5,
    2 => 10 + 10,
    3 => 15 + 15,
    _ => 0,
  };
  print(bonus);
}"#,
        ["30"]
    };

    switch_expr_yields_int_product_in_arm => {
        r#"void main() {
  var mult = switch (4) {
    2 => 3 * 4,
    4 => 5 * 6,
    _ => 0,
  };
  print(mult);
}"#,
        ["30"]
    };

    switch_expr_yields_int_from_nested_addition => {
        r#"void main() {
  var base = 2;
  var total = switch (base) {
    1 => 1 + 2 + 3,
    2 => 4 + 5 + 6,
    _ => 0,
  };
  print(total);
}"#,
        ["15"]
    };

    switch_expr_yields_int_used_in_multiplication => {
        r#"void main() {
  var factor = switch (2) {
    1 => 3,
    2 => 4,
    _ => 1,
  };
  print(factor * 5);
}"#,
        ["20"]
    };

    switch_expr_yields_int_returned_from_function => {
        r#"int scoreFor(int code) {
  return switch (code) {
    1 => 100,
    2 => 200,
    _ => 0,
  };
}
void main() {
  print(scoreFor(2));
}"#,
        ["200"]
    };

    switch_expr_yields_int_passed_as_argument => {
        r#"int doubleIt(int n) {
  return n * 2;
}
void main() {
  var code = 3;
  print(doubleIt(switch (code) {
    1 => 5,
    2 => 10,
    3 => 15,
    _ => 0,
  }));
}"#,
        ["30"]
    };

    switch_expr_yields_int_in_list_element => {
        r#"void main() {
  var code = 2;
  var list = [
    switch (code) {
      1 => 10,
      2 => 20,
      _ => 0,
    },
    99,
  ];
  print(list[0]);
  print(list[1]);
}"#,
        ["20", "99"]
    };

    switch_expr_yields_int_in_map_value => {
        r#"void main() {
  var key = 1;
  var map = {
    key: switch (key) {
      1 => 7,
      2 => 8,
      _ => 0,
    },
  };
  print(map[1]);
}"#,
        ["7"]
    };

    switch_expr_yields_string_for_dart_token => {
        r#"void main() {
  var lang = 'dart';
  var label = switch (lang) {
    'dart' => 'primary',
    'java' => 'legacy',
    _ => 'unknown',
  };
  print(label);
}"#,
        ["primary"]
    };

    switch_expr_yields_string_for_java_token => {
        r#"void main() {
  var lang = 'java';
  var label = switch (lang) {
    'dart' => 'primary',
    'java' => 'legacy',
    _ => 'unknown',
  };
  print(label);
}"#,
        ["legacy"]
    };

    switch_expr_yields_string_wildcard => {
        r#"void main() {
  var lang = 'kotlin';
  var label = switch (lang) {
    'dart' => 'primary',
    'java' => 'legacy',
    _ => 'unknown',
  };
  print(label);
}"#,
        ["unknown"]
    };

    switch_expr_yields_string_empty_literal => {
        r#"void main() {
  var s = switch ('') {
    '' => 'empty',
    _ => 'nonempty',
  };
  print(s);
}"#,
        ["empty"]
    };

    switch_expr_yields_string_concatenation_in_arm => {
        r#"void main() {
  var role = 'admin';
  var msg = switch (role) {
    'admin' => 'Hello ' + 'Admin',
    'user' => 'Hello ' + 'User',
    _ => 'Hello Guest',
  };
  print(msg);
}"#,
        ["Hello Admin"]
    };

    switch_expr_yields_string_interpolation_in_arm => {
        r#"void main() {
  var name = 'Ann';
  var greeting = switch (name.length) {
    3 => 'Hi $name',
    4 => 'Hello $name',
    _ => 'Hey $name',
  };
  print(greeting);
}"#,
        ["Hi Ann"]
    };

    switch_expr_yields_string_returned_from_function => {
        r#"String labelFor(int n) {
  return switch (n) {
    1 => 'one',
    2 => 'two',
    _ => 'many',
  };
}
void main() {
  print(labelFor(2));
}"#,
        ["two"]
    };

    switch_expr_yields_string_uppercase_in_arm => {
        r#"void main() {
  var mode = 'run';
  var banner = switch (mode) {
    'run' => 'RUNNING',
    'stop' => 'STOPPED',
    _ => 'IDLE',
  };
  print(banner);
}"#,
        ["RUNNING"]
    };

    switch_expr_nested_yields_int_inner_match => {
        r#"void main() {
  var outer = 1;
  var inner = 2;
  var result = switch (outer) {
    1 => switch (inner) {
      2 => 99,
      _ => 1,
    },
    _ => 0,
  };
  print(result);
}"#,
        ["99"]
    };

    switch_expr_nested_yields_int_outer_miss => {
        r#"void main() {
  var outer = 9;
  var inner = 2;
  var result = switch (outer) {
    1 => switch (inner) {
      2 => 99,
      _ => 1,
    },
    _ => 77,
  };
  print(result);
}"#,
        ["77"]
    };

    switch_expr_nested_yields_string_both_levels => {
        r#"void main() {
  var a = 'x';
  var b = 'y';
  var label = switch (a) {
    'x' => switch (b) {
      'y' => 'xy',
      _ => 'x-other',
    },
    _ => 'other',
  };
  print(label);
}"#,
        ["xy"]
    };

    switch_expr_nested_yields_string_inner_wildcard => {
        r#"void main() {
  var a = 'x';
  var b = 'z';
  var label = switch (a) {
    'x' => switch (b) {
      'y' => 'xy',
      _ => 'x-other',
    },
    _ => 'other',
  };
  print(label);
}"#,
        ["x-other"]
    };

    switch_expr_triple_nested_yields_int => {
        r#"void main() {
  var a = 1;
  var b = 2;
  var c = 3;
  var n = switch (a) {
    1 => switch (b) {
      2 => switch (c) {
        3 => 1000,
        _ => 100,
      },
      _ => 10,
    },
    _ => 0,
  };
  print(n);
}"#,
        ["1000"]
    };

    switch_expr_nested_in_return_statement => {
        r#"String pick(int a, int b) {
  return switch (a) {
    1 => switch (b) {
      2 => 'one-two',
      _ => 'one-other',
    },
    _ => 'other',
  };
}
void main() {
  print(pick(1, 2));
}"#,
        ["one-two"]
    };

    switch_expr_bool_yields_int_true_arm => {
        r#"void main() {
  var flag = true;
  var n = switch (flag) {
    true => 1,
    false => 0,
  };
  print(n);
}"#,
        ["1"]
    };

    switch_expr_bool_yields_int_false_arm => {
        r#"void main() {
  var flag = false;
  var n = switch (flag) {
    true => 1,
    false => 0,
  };
  print(n);
}"#,
        ["0"]
    };

    switch_expr_bool_yields_string => {
        r#"void main() {
  var ok = true;
  var msg = switch (ok) {
    true => 'yes',
    false => 'no',
  };
  print(msg);
}"#,
        ["yes"]
    };

    switch_expr_on_arithmetic_yields_int => {
        r#"void main() {
  var n = switch (3 + 4) {
    7 => 70,
    8 => 80,
    _ => 0,
  };
  print(n);
}"#,
        ["70"]
    };

    switch_expr_on_string_length_yields_int => {
        r#"void main() {
  var len = switch ('dart'.length) {
    3 => 30,
    4 => 40,
    _ => 0,
  };
  print(len);
}"#,
        ["40"]
    };

    switch_expr_two_sequential_int_yields => {
        r#"void main() {
  var a = switch (1) { 1 => 10, _ => 0 };
  var b = switch (2) { 2 => 20, _ => 0 };
  print(a);
  print(b);
}"#,
        ["10", "20"]
    };

    switch_expr_int_yields_used_in_subtraction => {
        r#"void main() {
  var base = 50;
  var delta = switch (3) {
    1 => 5,
    2 => 10,
    3 => 15,
    _ => 0,
  };
  print(base - delta);
}"#,
        ["35"]
    };

    switch_expr_int_yields_in_for_loop_accumulator => {
        r#"void main() {
  var sum = 0;
  for (var i = 1; i <= 3; i++) {
    sum += switch (i) {
      1 => 10,
      2 => 20,
      3 => 30,
      _ => 0,
    };
  }
  print(sum);
}"#,
        ["60"]
    };

    switch_expr_string_yields_in_conditional_print => {
        r#"void main() {
  var tier = 'gold';
  print(switch (tier) {
    'gold' => 'premium',
    'silver' => 'standard',
    _ => 'basic',
  });
}"#,
        ["premium"]
    };

    switch_expr_int_or_pattern_yields_int => {
        r#"void main() {
  var n = switch (2) {
    1 || 2 || 3 => 100,
    _ => 0,
  };
  print(n);
}"#,
        ["100"]
    };

    switch_expr_int_or_pattern_yields_string => {
        r#"void main() {
  var s = switch (5) {
    1 || 2 || 3 => 'small',
    4 || 5 || 6 => 'medium',
    _ => 'large',
  };
  print(s);
}"#,
        ["medium"]
    };

    switch_expr_string_or_pattern_yields_string => {
        r#"void main() {
  var day = 'Sat';
  var kind = switch (day) {
    'Sat' || 'Sun' => 'weekend',
    _ => 'weekday',
  };
  print(kind);
}"#,
        ["weekend"]
    };

    switch_expr_string_or_pattern_weekday_yields => {
        r#"void main() {
  var day = 'Wed';
  var kind = switch (day) {
    'Sat' || 'Sun' => 'weekend',
    _ => 'weekday',
  };
  print(kind);
}"#,
        ["weekday"]
    };

    switch_expr_int_relational_yields_int_small => {
        r#"void main() {
  var n = switch (4) {
    < 0 => -1,
    >= 0 && < 10 => 1,
    _ => 99,
  };
  print(n);
}"#,
        ["1"]
    };

    switch_expr_int_relational_yields_int_large => {
        r#"void main() {
  var n = switch (100) {
    < 0 => -1,
    >= 0 && < 10 => 1,
    _ => 99,
  };
  print(n);
}"#,
        ["99"]
    };

    switch_expr_int_relational_yields_string_negative => {
        r#"void main() {
  var label = switch (-1) {
    < 0 => 'below',
    0 => 'zero',
    _ => 'above',
  };
  print(label);
}"#,
        ["below"]
    };

    switch_expr_assign_to_typed_int_variable => {
        r#"void main() {
  int value = switch (2) {
    1 => 10,
    2 => 20,
    _ => 0,
  };
  print(value);
}"#,
        ["20"]
    };

    switch_expr_assign_to_typed_string_variable => {
        r#"void main() {
  String tag = switch ('b') {
    'a' => 'alpha',
    'b' => 'beta',
    _ => 'other',
  };
  print(tag);
}"#,
        ["beta"]
    };

    switch_expr_yields_int_from_variable_selector => {
        r#"void main() {
  var code = 4;
  var value = switch (code) {
    4 => 44,
    5 => 55,
    _ => 0,
  };
  print(value);
}"#,
        ["44"]
    };

    switch_expr_yields_string_from_variable_selector => {
        r#"void main() {
  var key = 'east';
  var dir = switch (key) {
    'north' => 'N',
    'south' => 'S',
    'east' => 'E',
    'west' => 'W',
    _ => '?',
  };
  print(dir);
}"#,
        ["E"]
    };

    switch_expr_nested_yields_int_in_addition => {
        r#"void main() {
  var a = 1;
  var b = 2;
  var total = switch (a) {
    1 => switch (b) {
      2 => 10,
      _ => 1,
    },
    _ => 0,
  } + switch (b) {
    2 => 20,
    _ => 0,
  };
  print(total);
}"#,
        ["30"]
    };

    switch_expr_int_yields_in_string_concat => {
        r#"void main() {
  var code = 7;
  var text = 'code-' + switch (code) {
    7 => 'seven',
    8 => 'eight',
    _ => 'other',
  }.toString();
  print(text);
}"#,
        ["code-seven"]
    };

    switch_expr_int_yields_max_of_two_branches => {
        r#"void main() {
  var pick = switch (2) {
    1 => 3,
    2 => 9,
    _ => 0,
  };
  var other = switch (1) {
    1 => 5,
    _ => 0,
  };
  print(pick > other ? pick : other);
}"#,
        ["9"]
    };
}
