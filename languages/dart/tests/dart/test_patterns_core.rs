//! Dart 3 patterns: switch expressions on int/String/List/Record, if-case,
//! wildcard `_`, logical-or patterns, guarded `when` clauses, and destructuring.

dart_cases! {
    switch_expr_int_matches_literal_one => {
        r#"void main() {
  var n = 1;
  print(switch (n) {
    1 => 'one',
    2 => 'two',
    _ => 'other' });
}"#,
        ["one"]
    };

    switch_expr_int_matches_literal_two => {
        r#"void main() {
  var n = 2;
  print(switch (n) {
    1 => 'one',
    2 => 'two',
    _ => 'other' });
}"#,
        ["two"]
    };

    switch_expr_int_falls_through_to_wildcard => {
        r#"void main() {
  var n = 99;
  print(switch (n) {
    1 => 'one',
    2 => 'two',
    _ => 'other' });
}"#,
        ["other"]
    };

    switch_expr_int_zero_literal => {
        r#"void main() {
  print(switch (0) {
    0 => 'zero',
    _ => 'nonzero' });
}"#,
        ["zero"]
    };

    switch_expr_int_negative_literal => {
        r#"void main() {
  print(switch (-3) {
    -3 => 'neg-three',
    _ => 'other' });
}"#,
        ["neg-three"]
    };

    switch_expr_int_relational_less_than => {
        r#"void main() {
  print(switch (4) {
    < 0 => 'negative',
    >= 0 && < 10 => 'small',
    _ => 'large' });
}"#,
        ["small"]
    };

    switch_expr_int_relational_negative_branch => {
        r#"void main() {
  print(switch (-5) {
    < 0 => 'negative',
    0 => 'zero',
    _ => 'positive' });
}"#,
        ["negative"]
    };

    switch_expr_int_relational_large_branch => {
        r#"void main() {
  print(switch (100) {
    < 0 => 'negative',
    >= 0 && < 10 => 'small',
    _ => 'large' });
}"#,
        ["large"]
    };

    switch_expr_int_equality_exact => {
        r#"void main() {
  print(switch (42) {
    41 => 'miss',
    42 => 'hit',
    43 => 'miss2' });
}"#,
        ["hit"]
    };

    switch_expr_int_nested_switch => {
        r#"void main() {
  var outer = 1;
  var inner = 2;
  print(switch (outer) {
    1 => switch (inner) {
      2 => 'nested',
      _ => 'inner-other' },
    _ => 'outer-other' });
}"#,
        ["nested"]
    };

    switch_expr_int_result_used_in_addition => {
        r#"void main() {
  var code = 3;
  var label = switch (code) {
    1 => 10,
    2 => 20,
    3 => 30,
    _ => 0 };
  print(label + 5);
}"#,
        ["35"]
    };

    switch_expr_string_matches_literal => {
        r#"void main() {
  var s = 'dart';
  print(switch (s) {
    'dart' => 'match',
    'java' => 'miss',
    _ => 'other' });
}"#,
        ["match"]
    };

    switch_expr_string_no_match_wildcard => {
        r#"void main() {
  var s = 'kotlin';
  print(switch (s) {
    'dart' => 'match',
    'java' => 'miss',
    _ => 'other' });
}"#,
        ["other"]
    };

    switch_expr_string_empty_literal => {
        r#"void main() {
  print(switch ('') {
    '' => 'empty',
    _ => 'nonempty' });
}"#,
        ["empty"]
    };

    switch_expr_string_single_char => {
        r#"void main() {
  print(switch ('x') {
    'x' => 'ex',
    'y' => 'why',
    _ => 'other' });
}"#,
        ["ex"]
    };

    switch_expr_string_multi_word_token => {
        r#"void main() {
  print(switch ('hello world') {
    'hello world' => 'greeting',
    'bye' => 'farewell',
    _ => 'unknown' });
}"#,
        ["greeting"]
    };

    switch_expr_string_case_sensitive_miss => {
        r#"void main() {
  print(switch ('Dart') {
    'dart' => 'lower',
    'DART' => 'upper',
    _ => 'mixed' });
}"#,
        ["mixed"]
    };

    switch_expr_string_or_pattern_weekend => {
        r#"void main() {
  var day = 'Sat';
  print(switch (day) {
    'Sat' || 'Sun' => 'weekend',
    _ => 'weekday' });
}"#,
        ["weekend"]
    };

    switch_expr_string_or_pattern_weekday => {
        r#"void main() {
  var day = 'Mon';
  print(switch (day) {
    'Sat' || 'Sun' => 'weekend',
    _ => 'weekday' });
}"#,
        ["weekday"]
    };

    switch_expr_string_three_way_or => {
        r#"void main() {
  print(switch ('b') {
    'a' || 'b' || 'c' => 'abc',
    _ => 'other' });
}"#,
        ["abc"]
    };

    switch_expr_string_or_misses_all => {
        r#"void main() {
  print(switch ('z') {
    'a' || 'b' || 'c' => 'abc',
    _ => 'other' });
}"#,
        ["other"]
    };

    switch_expr_wildcard_catches_unmatched_int => {
        r#"void main() {
  print(switch (7) {
    1 => 'one',
    _ => 'catch-all' });
}"#,
        ["catch-all"]
    };

    switch_expr_wildcard_only_arm => {
        r#"void main() {
  print(switch (999) {
    _ => 'always' });
}"#,
        ["always"]
    };

    switch_expr_wildcard_after_specific_arms => {
        r#"void main() {
  print(switch ('x') {
    'a' => 'alpha',
    'b' => 'beta',
    _ => 'rest' });
}"#,
        ["rest"]
    };

    switch_expr_wildcard_with_null_subject => {
        r#"void main() {
  print(switch (null) {
    null => 'is-null',
    _ => 'not-null' });
}"#,
        ["is-null"]
    };

    switch_expr_wildcard_bool_true_arm => {
        r#"void main() {
  print(switch (true) {
    true => 'yes',
    false => 'no' });
}"#,
        ["yes"]
    };

    switch_expr_int_or_pattern_small_values => {
        r#"void main() {
  print(switch (2) {
    1 || 2 || 3 => 'small',
    _ => 'big' });
}"#,
        ["small"]
    };

    switch_expr_int_or_pattern_first_arm => {
        r#"void main() {
  print(switch (1) {
    1 || 2 || 3 => 'small',
    _ => 'big' });
}"#,
        ["small"]
    };

    switch_expr_int_or_pattern_third_arm => {
        r#"void main() {
  print(switch (3) {
    1 || 2 || 3 => 'small',
    _ => 'big' });
}"#,
        ["small"]
    };

    switch_expr_int_or_pattern_miss => {
        r#"void main() {
  print(switch (10) {
    1 || 2 || 3 => 'small',
    _ => 'big' });
}"#,
        ["big"]
    };

    switch_expr_int_or_and_combined => {
        r#"void main() {
  print(switch (5) {
    1 || 2 => 'tiny',
    3 || 4 || 5 => 'mid',
    _ => 'large' });
}"#,
        ["mid"]
    };

    switch_expr_int_or_two_groups => {
        r#"void main() {
  print(switch (8) {
    1 || 2 => 'pair-a',
    7 || 8 || 9 => 'pair-b',
    _ => 'neither' });
}"#,
        ["pair-b"]
    };

    switch_expr_int_or_with_zero => {
        r#"void main() {
  print(switch (0) {
    0 || 1 => 'zero-or-one',
    _ => 'other' });
}"#,
        ["zero-or-one"]
    };

    switch_expr_int_or_negative_values => {
        r#"void main() {
  print(switch (-1) {
    -1 || -2 => 'negative-pair',
    _ => 'other' });
}"#,
        ["negative-pair"]
    };

    switch_expr_list_empty_pattern => {
        r#"void main() {
  var xs = <int>[];
  print(switch (xs) {
    [] => 'empty',
    _ => 'nonempty' });
}"#,
        ["empty"]
    };

    switch_expr_list_single_element_pattern => {
        r#"void main() {
  var xs = [7];
  print(switch (xs) {
    [] => 'empty',
    [var a] => 'single',
    _ => 'multi' });
}"#,
        ["single"]
    };

    switch_expr_list_two_element_pattern => {
        r#"void main() {
  var xs = [1, 2];
  print(switch (xs) {
    [] => 'empty',
    [var a, var b] => 'pair',
    _ => 'other' });
}"#,
        ["pair"]
    };

    switch_expr_list_three_element_pattern => {
        r#"void main() {
  var xs = [10, 20, 30];
  print(switch (xs) {
    [var a, var b, var c] => 'triple',
    _ => 'other' });
}"#,
        ["triple"]
    };

    switch_expr_list_destructure_binds_values => {
        r#"void main() {
  var xs = [3, 4];
  var sum = switch (xs) {
    [var a, var b] => a + b,
    _ => 0 };
  print(sum);
}"#,
        ["7"]
    };

    switch_expr_list_first_rest_pattern => {
        r#"void main() {
  var xs = [1, 2, 3, 4];
  print(switch (xs) {
    [var head, ...var tail] => head,
    _ => -1 });
}"#,
        ["1"]
    };

    switch_expr_list_rest_length_via_destructure => {
        r#"void main() {
  var xs = [9, 8, 7];
  var count = switch (xs) {
    [var first, ...var rest] => rest.length + 1,
    _ => 0 };
  print(count);
}"#,
        ["3"]
    };

    switch_expr_list_string_elements => {
        r#"void main() {
  var xs = ['a', 'b'];
  print(switch (xs) {
    ['a', 'b'] => 'ab',
    _ => 'other' });
}"#,
        ["ab"]
    };

    switch_expr_list_mixed_length_wildcard => {
        r#"void main() {
  var xs = [1, 2, 3, 4, 5];
  print(switch (xs) {
    [] => 'empty',
    [var a] => 'one',
    [var a, var b] => 'two',
    _ => 'many' });
}"#,
        ["many"]
    };

    switch_expr_list_nested_list_pattern => {
        r#"void main() {
  var xs = [[1, 2], [3, 4]];
  print(switch (xs) {
    [var a, var b] => 'pair-of-lists',
    _ => 'other' });
}"#,
        ["pair-of-lists"]
    };

    switch_expr_list_wildcard_element => {
        r#"void main() {
  var xs = [5, 9];
  print(switch (xs) {
    [var _, var y] => y,
    _ => 0 });
}"#,
        ["9"]
    };

    switch_expr_list_constant_first_slot => {
        r#"void main() {
  var xs = [0, 99];
  print(switch (xs) {
    [0, var n] => n,
    _ => -1 });
}"#,
        ["99"]
    };

    switch_expr_list_or_two_shapes => {
        r#"void main() {
  var xs = [1];
  print(switch (xs) {
    [] || [var _] => 'empty-or-one',
    _ => 'other' });
}"#,
        ["empty-or-one"]
    };

    switch_expr_record_positional_literal_match => {
        r#"void main() {
  var p = (0, 0);
  print(switch (p) {
    (0, 0) => 'origin',
    _ => 'elsewhere' });
}"#,
        ["origin"]
    };

    switch_expr_record_positional_destructure => {
        r#"void main() {
  var p = (3, 4);
  var total = switch (p) {
    (var x, var y) => x + y,
    _ => 0 };
  print(total);
}"#,
        ["7"]
    };

    switch_expr_record_named_field_match => {
        r#"void main() {
  var u = (name: 'Ada', id: 42);
  print(switch (u) {
    (name: 'Ada', id: var n) => n,
    _ => 0 });
}"#,
        ["42"]
    };

    switch_expr_record_named_destructure => {
        r#"void main() {
  var u = (name: 'Bob', score: 10);
  print(switch (u) {
    (name: var n, score: var s) => s,
    _ => 0 });
}"#,
        ["10"]
    };

    switch_expr_record_mixed_positional_named => {
        r#"void main() {
  var e = (1, label: 'one');
  print(switch (e) {
    (var n, label: var lbl) => lbl,
    _ => 'none' });
}"#,
        ["one"]
    };

    switch_expr_record_wildcard_field => {
        r#"void main() {
  var p = (7, 8);
  print(switch (p) {
    (var _, var y) => y,
    _ => 0 });
}"#,
        ["8"]
    };

    switch_expr_record_three_positional => {
        r#"void main() {
  var rgb = (1, 2, 3);
  print(switch (rgb) {
    (var r, var g, var b) => r + g + b,
    _ => 0 });
}"#,
        ["6"]
    };

    switch_expr_record_or_two_shapes => {
        r#"void main() {
  var p = (1, 2);
  print(switch (p) {
    (0, 0) || (1, 2) => 'special',
    _ => 'generic' });
}"#,
        ["special"]
    };

    switch_expr_record_no_match_wildcard => {
        r#"void main() {
  var p = (9, 9);
  print(switch (p) {
    (0, 0) => 'origin',
    _ => 'other' });
}"#,
        ["other"]
    };

    switch_expr_when_positive_guard => {
        r#"void main() {
  print(switch (5) {
    var n when n > 0 => 'positive',
    var n when n < 0 => 'negative',
    _ => 'zero' });
}"#,
        ["positive"]
    };

    switch_expr_when_negative_guard => {
        r#"void main() {
  print(switch (-2) {
    var n when n > 0 => 'positive',
    var n when n < 0 => 'negative',
    _ => 'zero' });
}"#,
        ["negative"]
    };

    switch_expr_when_zero_falls_to_wildcard => {
        r#"void main() {
  print(switch (0) {
    var n when n > 0 => 'positive',
    var n when n < 0 => 'negative',
    _ => 'zero' });
}"#,
        ["zero"]
    };

    switch_expr_when_string_length_guard => {
        r#"void main() {
  print(switch ('hello') {
    var s when s.length == 5 => 'five-chars',
    _ => 'other' });
}"#,
        ["five-chars"]
    };

    switch_expr_when_list_sum_guard => {
        r#"void main() {
  var xs = [1, 2, 3];
  print(switch (xs) {
    [var a, var b, var c] when a + b + c == 6 => 'sum-six',
    _ => 'other' });
}"#,
        ["sum-six"]
    };

    switch_expr_when_record_score_threshold => {
        r#"void main() {
  var r = (name: 'Ann', score: 95);
  print(switch (r) {
    (name: var n, score: var s) when s >= 90 => 'A',
    (name: var n, score: var s) when s >= 80 => 'B',
    _ => 'C' });
}"#,
        ["A"]
    };

    switch_expr_when_int_or_with_guard => {
        r#"void main() {
  print(switch (4) {
    1 || 2 || 3 => 'small',
    var n when n >= 4 && n < 10 => 'mid',
    _ => 'large' });
}"#,
        ["mid"]
    };

    switch_expr_when_even_odd_guard => {
        r#"void main() {
  print(switch (6) {
    var n when n % 2 == 0 => 'even',
    var n when n % 2 == 1 => 'odd',
    _ => 'unknown' });
}"#,
        ["even"]
    };

    switch_expr_when_destructure_list_guard => {
        r#"void main() {
  var xs = [2, 4];
  print(switch (xs) {
    [var a, var b] when a * b == 8 => 'product-eight',
    _ => 'other' });
}"#,
        ["product-eight"]
    };

    switch_expr_when_record_named_guard => {
        r#"void main() {
  var u = (role: 'admin', level: 3);
  print(switch (u) {
    (role: 'admin', level: var lv) when lv >= 2 => 'elevated',
    _ => 'standard' });
}"#,
        ["elevated"]
    };

    if_case_int_literal_match => {
        r#"void main() {
  var n = 42;
  if (n case 42) {
    print('match');
  } else {
    print('miss');
  }
}"#,
        ["match"]
    };

    if_case_int_literal_miss => {
        r#"void main() {
  var n = 7;
  if (n case 42) {
    print('match');
  } else {
    print('miss');
  }
}"#,
        ["miss"]
    };

    if_case_string_literal_match => {
        r#"void main() {
  var s = 'dart';
  if (s case 'dart') {
    print('yes');
  } else {
    print('no');
  }
}"#,
        ["yes"]
    };

    if_case_list_destructure_two_elements => {
        r#"void main() {
  var xs = [10, 20];
  if (xs case [var a, var b]) {
    print(a);
    print(b);
  }
}"#,
        ["10", "20"]
    };

    if_case_list_empty_pattern => {
        r#"void main() {
  var xs = <int>[];
  if (xs case []) {
    print('empty');
  } else {
    print('nonempty');
  }
}"#,
        ["empty"]
    };

    if_case_record_positional_destructure => {
        r#"void main() {
  var p = (3, 5);
  if (p case (var x, var y)) {
    print(x + y);
  }
}"#,
        ["8"]
    };

    if_case_record_named_destructure => {
        r#"void main() {
  var u = (name: 'Eve', age: 30);
  if (u case (name: var n, age: var a)) {
    print(n);
    print(a);
  }
}"#,
        ["Eve", "30"]
    };

    if_case_wildcard_list_second_slot => {
        r#"void main() {
  var xs = [1, 99];
  if (xs case [var _, var tail]) {
    print(tail);
  }
}"#,
        ["99"]
    };

    if_case_when_positive_guard => {
        r#"void main() {
  var n = 12;
  if (n case var x when x > 10) {
    print('big');
  } else {
    print('small');
  }
}"#,
        ["big"]
    };

    if_case_when_guard_misses_else => {
        r#"void main() {
  var n = 3;
  if (n case var x when x > 10) {
    print('big');
  } else {
    print('small');
  }
}"#,
        ["small"]
    };

    if_case_or_pattern_int => {
        r#"void main() {
  var code = 2;
  if (code case 1 || 2 || 3) {
    print('small');
  } else {
    print('large');
  }
}"#,
        ["small"]
    };

    destructuring_var_positional_record => {
        r#"void main() {
  var (x, y) = (6, 7);
  print(x);
  print(y);
}"#,
        ["6", "7"]
    };

    destructuring_var_named_record => {
        r#"void main() {
  var (name: n, count: c) = (name: 'pi', count: 3);
  print(n);
  print(c);
}"#,
        ["pi", "3"]
    };

    destructuring_var_list_pattern => {
        r#"void main() {
  var [first, second] = [100, 200];
  print(first);
  print(second);
}"#,
        ["100", "200"]
    };

    destructuring_in_switch_binds_record_fields => {
        r#"void main() {
  var item = (id: 5, qty: 2);
  switch (item) {
    case (id: var i, qty: var q):
      print(i * q);
      break;
    default:
      print(0);
  }
}"#,
        ["10"]
    };
}
