//! List and map patterns only: [a,b], rest ..., map key patterns, and guarded when clauses.

dart_cases! {
    switch_list_empty_matches_empty_arm => {
        r#"void main() {
  var xs = <int>[];
  print(switch (xs) {
    [] => 'empty',
    _ => 'other',
  });
}"#,
        ["empty"]
    };

    switch_list_single_var_binds_element => {
        r#"void main() {
  var xs = [42];
  print(switch (xs) {
    [var n] => n,
    _ => -1,
  });
}"#,
        ["42"]
    };

    switch_list_pair_destructure => {
        r#"void main() {
  var xs = [3, 4];
  print(switch (xs) {
    [var a, var b] => a + b,
    _ => 0,
  });
}"#,
        ["7"]
    };

    switch_list_triple_destructure => {
        r#"void main() {
  var xs = [1, 2, 3];
  print(switch (xs) {
    [var a, var b, var c] => a + b + c,
    _ => 0,
  });
}"#,
        ["6"]
    };

    switch_list_rest_captures_tail => {
        r#"void main() {
  var xs = [1, 2, 3, 4];
  print(switch (xs) {
    [var head, ...var tail] => tail.length,
    _ => -1,
  });
}"#,
        ["3"]
    };

    switch_list_rest_head_value => {
        r#"void main() {
  var xs = [9, 8, 7];
  print(switch (xs) {
    [var first, ...var _] => first,
    _ => 0,
  });
}"#,
        ["9"]
    };

    switch_list_rest_on_two_element_list => {
        r#"void main() {
  var xs = [5, 6];
  print(switch (xs) {
    [var a, ...var rest] => rest.length + a,
    _ => 0,
  });
}"#,
        ["6"]
    };

    switch_list_rest_on_single_element => {
        r#"void main() {
  var xs = [99];
  print(switch (xs) {
    [var a, ...var rest] => rest.isEmpty,
    _ => false,
  });
}"#,
        ["true"]
    };

    switch_list_constant_first_slot => {
        r#"void main() {
  var xs = [0, 15];
  print(switch (xs) {
    [0, var n] => n,
    _ => -1,
  });
}"#,
        ["15"]
    };

    switch_list_wildcard_second_slot => {
        r#"void main() {
  var xs = [1, 8];
  print(switch (xs) {
    [var _, var y] => y,
    _ => 0,
  });
}"#,
        ["8"]
    };

    switch_list_string_literal_elements => {
        r#"void main() {
  var xs = ['a', 'b'];
  print(switch (xs) {
    ['a', 'b'] => 'match',
    _ => 'miss',
  });
}"#,
        ["match"]
    };

    switch_list_length_branch_many => {
        r#"void main() {
  var xs = [1, 2, 3, 4, 5];
  print(switch (xs) {
    [] => 'empty',
    [var _] => 'one',
    [var _, var __] => 'two',
    _ => 'many',
  });
}"#,
        ["many"]
    };

    switch_list_nested_inner_pattern => {
        r#"void main() {
  var xs = [[1, 2], [3, 4]];
  print(switch (xs) {
    [var a, var b] => a.length + b.length,
    _ => 0,
  });
}"#,
        ["4"]
    };

    switch_list_or_empty_or_one => {
        r#"void main() {
  var xs = <int>[];
  print(switch (xs) {
    [] || [var _] => 'small',
    _ => 'big',
  });
}"#,
        ["small"]
    };

    switch_list_when_sum_guard_passes => {
        r#"void main() {
  var xs = [2, 3];
  print(switch (xs) {
    [var a, var b] when a + b == 5 => 'five',
    _ => 'other',
  });
}"#,
        ["five"]
    };

    switch_list_when_sum_guard_fails => {
        r#"void main() {
  var xs = [2, 4];
  print(switch (xs) {
    [var a, var b] when a + b == 5 => 'five',
    _ => 'other',
  });
}"#,
        ["other"]
    };

    switch_list_when_rest_length_guard => {
        r#"void main() {
  var xs = [1, 2, 3, 4];
  print(switch (xs) {
    [var _, ...var tail] when tail.length == 3 => 'three-tail',
    _ => 'other',
  });
}"#,
        ["three-tail"]
    };

    if_case_list_pair_destructure => {
        r#"void main() {
  var xs = [10, 20];
  if (xs case [var a, var b]) {
    print(a + b);
  } else {
    print(0);
  }
}"#,
        ["30"]
    };

    if_case_list_empty_pattern => {
        r#"void main() {
  var xs = <int>[];
  if (xs case []) {
    print('empty');
  } else {
    print('other');
  }
}"#,
        ["empty"]
    };

    if_case_list_rest_pattern => {
        r#"void main() {
  var xs = [1, 2, 3];
  if (xs case [var h, ...var t]) {
    print(h);
    print(t.length);
  } else {
    print(-1);
  }
}"#,
        ["1", "2"]
    };

    if_case_list_when_positive_guard => {
        r#"void main() {
  var xs = [4, 5];
  if (xs case [var a, var b] when a < b) {
    print('asc');
  } else {
    print('no');
  }
}"#,
        ["asc"]
    };

    switch_map_single_key_var_value => {
        r#"void main() {
  var m = {'a': 1};
  print(switch (m) {
    {'a': var v} => v,
    _ => -1,
  });
}"#,
        ["1"]
    };

    switch_map_two_key_pattern => {
        r#"void main() {
  var m = {'x': 10, 'y': 20};
  print(switch (m) {
    {'x': var a, 'y': var b} => a + b,
    _ => 0,
  });
}"#,
        ["30"]
    };

    switch_map_literal_key_match => {
        r#"void main() {
  var m = {'mode': 'on', 'level': 3};
  print(switch (m) {
    {'mode': 'on', 'level': var n} => n,
    _ => 0,
  });
}"#,
        ["3"]
    };

    switch_map_wildcard_value_slot => {
        r#"void main() {
  var m = {'k': 99};
  print(switch (m) {
    {'k': var _} => 'found',
    _ => 'miss',
  });
}"#,
        ["found"]
    };

    switch_map_when_value_guard => {
        r#"void main() {
  var m = {'score': 85};
  print(switch (m) {
    {'score': var s} when s >= 80 => 'pass',
    _ => 'fail',
  });
}"#,
        ["pass"]
    };

    switch_map_when_value_guard_miss => {
        r#"void main() {
  var m = {'score': 55};
  print(switch (m) {
    {'score': var s} when s >= 80 => 'pass',
    _ => 'fail',
  });
}"#,
        ["fail"]
    };

    if_case_map_two_fields => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  if (m case {'a': var x, 'b': var y}) {
    print(x + y);
  } else {
    print(0);
  }
}"#,
        ["3"]
    };

    if_case_map_when_key_present => {
        r#"void main() {
  var m = {'id': 7};
  if (m case {'id': var n} when n > 0) {
    print('ok');
  } else {
    print('no');
  }
}"#,
        ["ok"]
    };

    switch_list_four_element_destructure => {
        r#"void main() {
  var xs = [1, 2, 3, 4];
  print(switch (xs) {
    [var a, var b, var c, var d] => a + b + c + d,
    _ => 0,
  });
}"#,
        ["10"]
    };

    switch_list_rest_sum_tail => {
        r#"void main() {
  var xs = [2, 3, 4, 5];
  print(switch (xs) {
    [var first, ...var rest] => first + rest.fold(0, (a, b) => a + b),
    _ => 0,
  });
}"#,
        ["14"]
    };

    switch_list_negative_constant_match => {
        r#"void main() {
  var xs = [-1, 5];
  print(switch (xs) {
    [-1, var n] => n,
    _ => 0,
  });
}"#,
        ["5"]
    };

    switch_list_double_rest_guard_on_length => {
        r#"void main() {
  var xs = [1, 2, 3, 4, 5];
  print(switch (xs) {
    [var _, var _, ...var tail] when tail.length == 3 => 'ok',
    _ => 'no',
  });
}"#,
        ["ok"]
    };

    switch_map_three_field_destructure => {
        r#"void main() {
  var m = {'r': 1, 'g': 2, 'b': 3};
  print(switch (m) {
    {'r': var r, 'g': var g, 'b': var b} => r + g + b,
    _ => 0,
  });
}"#,
        ["6"]
    };

    switch_map_or_two_shapes => {
        r#"void main() {
  var m = {'type': 'a'};
  print(switch (m) {
    {'type': 'a'} || {'type': 'b'} => 'letter',
    _ => 'other',
  });
}"#,
        ["letter"]
    };

    switch_map_miss_falls_to_wildcard => {
        r#"void main() {
  var m = {'x': 1};
  print(switch (m) {
    {'y': var _} => 'y',
    _ => 'fallback',
  });
}"#,
        ["fallback"]
    };

    if_case_map_single_entry => {
        r#"void main() {
  var m = {'only': 42};
  if (m case {'only': var v}) {
    print(v);
  } else {
    print(0);
  }
}"#,
        ["42"]
    };

    if_case_list_wildcard_first => {
        r#"void main() {
  var xs = [0, 7];
  if (xs case [var _, var last]) {
    print(last);
  } else {
    print(-1);
  }
}"#,
        ["7"]
    };

    switch_list_bool_elements => {
        r#"void main() {
  var xs = [true, false];
  print(switch (xs) {
    [true, var b] => b,
    _ => true,
  });
}"#,
        ["false"]
    };

    switch_map_bool_value_guard => {
        r#"void main() {
  var m = {'enabled': true};
  print(switch (m) {
    {'enabled': var flag} when flag => 'on',
    _ => 'off',
  });
}"#,
        ["on"]
    };

    switch_list_longer_rest_tail_join => {
        r#"void main() {
  var xs = [1, 2, 3, 4, 5, 6];
  print(switch (xs) {
    [var a, var b, ...var rest] => rest.length + a + b,
    _ => 0,
  });
}"#,
        ["7"]
    };

    switch_map_string_values_concat => {
        r#"void main() {
  var m = {'first': 'hello', 'second': 'world'};
  print(switch (m) {
    {'first': var a, 'second': var b} => a + b,
    _ => '',
  });
}"#,
        ["helloworld"]
    };

    switch_list_empty_rest_not_reached => {
        r#"void main() {
  var xs = [8];
  print(switch (xs) {
    [] => 'empty',
    [var x, ...var r] => r.length,
    _ => -1,
  });
}"#,
        ["0"]
    };

    if_case_map_when_guard_false => {
        r#"void main() {
  var m = {'n': 0};
  if (m case {'n': var v} when v > 0) {
    print('yes');
  } else {
    print('no');
  }
}"#,
        ["no"]
    };

    switch_list_nested_destructure_inner_pair => {
        r#"void main() {
  var xs = [[1, 2], [3, 4]];
  print(switch (xs) {
    [[var a, var b], [var c, var d]] => a + b + c + d,
    _ => 0,
  });
}"#,
        ["10"]
    };

    switch_map_int_keys_as_string_literals => {
        r#"void main() {
  var m = {'1': 10, '2': 20};
  print(switch (m) {
    {'1': var x, '2': var y} => x + y,
    _ => 0,
  });
}"#,
        ["30"]
    };

    switch_list_when_first_element_even => {
        r#"void main() {
  var xs = [4, 9];
  print(switch (xs) {
    [var a, var b] when a.isEven => b,
    _ => 0,
  });
}"#,
        ["9"]
    };

    switch_map_partial_key_set => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  print(switch (m) {
    {'a': var x, 'c': var z} => x + z,
    _ => 0,
  });
}"#,
        ["4"]
    };

    if_case_list_three_with_rest => {
        r#"void main() {
  var xs = [1, 2, 3, 4];
  if (xs case [var a, ...var rest]) {
    print(a);
    print(rest.join(','));
  } else {
    print('no');
  }
}"#,
        ["1", "2,3,4"]
    };

    switch_list_or_two_lengths => {
        r#"void main() {
  var xs = [1, 2];
  print(switch (xs) {
    [var _] || [var _, var __] => 'short',
    _ => 'long',
  });
}"#,
        ["short"]
    };

    switch_map_when_two_field_sum => {
        r#"void main() {
  var m = {'u': 2, 'v': 3};
  print(switch (m) {
    {'u': var u, 'v': var v} when u + v == 5 => 'sum5',
    _ => 'other',
  });
}"#,
        ["sum5"]
    };

    switch_list_double_element_string_pattern => {
        r#"void main() {
  var xs = ['go', 'dart'];
  print(switch (xs) {
    ['go', 'dart'] => 'both',
    _ => 'miss',
  });
}"#,
        ["both"]
    };

    switch_map_empty_map_wildcard => {
        r#"void main() {
  var m = <String, int>{};
  print(switch (m) {
    {} => 'empty-map',
    _ => 'nonempty',
  });
}"#,
        ["empty-map"]
    };

    if_case_map_rest_not_used_two_keys => {
        r#"void main() {
  var m = {'p': 1, 'q': 2, 'r': 3};
  if (m case {'p': var x, 'r': var z}) {
    print(x);
    print(z);
  } else {
    print(0);
  }
}"#,
        ["1", "3"]
    };

    switch_list_rest_preserves_order_sum => {
        r#"void main() {
  var xs = [10, 1, 2, 3];
  print(switch (xs) {
    [var h, ...var t] => h + t[0] + t[1] + t[2],
    _ => 0,
  });
}"#,
        ["16"]
    };

    switch_map_nested_value_via_var => {
        r#"void main() {
  var m = {'data': [1, 2, 3]};
  print(switch (m) {
    {'data': var list} => list.length,
    _ => 0,
  });
}"#,
        ["3"]
    };

    switch_list_guard_on_rest_nonempty => {
        r#"void main() {
  var xs = [5, 6, 7];
  print(switch (xs) {
    [var _, ...var tail] when tail.isNotEmpty => tail.first,
    _ => 0,
  });
}"#,
        ["6"]
    };

    switch_map_guard_on_string_value_length => {
        r#"void main() {
  var m = {'name': 'dart'};
  print(switch (m) {
    {'name': var s} when s.length == 4 => 'four',
    _ => 'other',
  });
}"#,
        ["four"]
    };
}
