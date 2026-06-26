//! break/continue in for, while, do-while, for-in, and nested labeled loops.

dart_cases! {
    for_break_stops_factorial_at_four => {
        r#"void main() {
  var fact = 1;
  for (var i = 1; i <= 8; i++) {
    fact *= i;
    if (i == 4) break;
  }
  print(fact);
}"#,
        ["24"]
    };

    for_continue_skips_multiples_of_four => {
        r#"void main() {
  for (var i = 1; i <= 8; i++) {
    if (i % 4 == 0) continue;
    print(i);
  }
}"#,
        ["1", "2", "3", "5", "6", "7", "8"]
    };

    for_break_on_first_iteration => {
        r#"void main() {
  var count = 0;
  for (var i = 0; i < 5; i++) {
    count++;
    break;
  }
  print(count);
}"#,
        ["1"]
    };

    for_continue_on_first_iteration_skips_body_rest => {
        r#"void main() {
  var printed = 0;
  for (var i = 0; i < 3; i++) {
    continue;
    printed++;
  }
  print(printed);
}"#,
        ["0"]
    };

    for_break_leaves_partial_product => {
        r#"void main() {
  var product = 1;
  for (var i = 1; i <= 6; i++) {
    if (i == 4) break;
    product *= i;
  }
  print(product);
}"#,
        ["6"]
    };

    for_continue_collects_only_primes_below_ten => {
        r#"void main() {
  var count = 0;
  for (var i = 2; i < 10; i++) {
    if (i == 4 || i == 6 || i == 8) continue;
    count++;
  }
  print(count);
}"#,
        ["4"]
    };

    while_break_after_three_prints => {
        r#"void main() {
  var i = 0;
  while (i < 10) {
    print(i);
    i++;
    if (i == 3) break;
  }
}"#,
        ["0", "1", "2"]
    };

    while_continue_skips_odd_values => {
        r#"void main() {
  var i = 0;
  while (i < 6) {
    i++;
    if (i % 2 == 1) continue;
    print(i);
  }
}"#,
        ["2", "4", "6"]
    };

    while_break_on_zero_guard => {
        r#"void main() {
  var n = 8;
  while (n > 0) {
    if (n == 3) break;
    n--;
  }
  print(n);
}"#,
        ["3"]
    };

    do_while_break_on_second_iteration => {
        r#"void main() {
  var i = 0;
  do {
    i++;
    if (i == 2) break;
    print(i);
  } while (i < 5);
}"#,
        ["1"]
    };

    do_while_continue_skips_printing_three => {
        r#"void main() {
  var i = 0;
  do {
    i++;
    if (i == 3) continue;
    print(i);
  } while (i < 5);
}"#,
        ["1", "2", "4", "5"]
    };

    do_while_continue_then_break => {
        r#"void main() {
  var i = 0;
  do {
    i++;
    if (i == 2) continue;
    if (i == 4) break;
    print(i);
  } while (i < 10);
}"#,
        ["1", "3"]
    };

    for_in_break_on_target_string_char => {
        r#"void main() {
  var hit = '';
  for (var ch in 'abcde') {
    if (ch == 'c') {
      hit = ch;
      break;
    }
  }
  print(hit);
}"#,
        ["c"]
    };

    for_in_continue_skips_zeros_in_list => {
        r#"void main() {
  var sum = 0;
  for (var x in [1, 0, 2, 0, 3]) {
    if (x == 0) continue;
    sum += x;
  }
  print(sum);
}"#,
        ["6"]
    };

    for_in_break_before_last_element => {
        r#"void main() {
  var count = 0;
  for (var x in [5, 6, 7, 8, 9]) {
    count++;
    if (x == 8) break;
  }
  print(count);
}"#,
        ["4"]
    };

    for_in_continue_filters_even_set_members => {
        r#"void main() {
  var sum = 0;
  for (var x in {1, 2, 3, 4}) {
    if (x % 2 == 0) continue;
    sum += x;
  }
  print(sum);
}"#,
        ["4"]
    };

    nested_for_break_inner_after_two_hits_per_row => {
        r#"void main() {
  var count = 0;
  for (var i = 0; i < 2; i++) {
    for (var j = 0; j < 4; j++) {
      count++;
      if (j == 1) break;
    }
  }
  print(count);
}"#,
        ["4"]
    };

    nested_for_continue_inner_skips_last_column => {
        r#"void main() {
  for (var i = 0; i < 2; i++) {
    for (var j = 0; j < 3; j++) {
      if (j == 2) continue;
      print('$i$j');
    }
  }
}"#,
        ["00", "01", "10", "11"]
    };

    nested_while_break_inner_preserves_outer => {
        r#"void main() {
  var outer = 0;
  var r = 0;
  while (r < 2) {
    var c = 0;
    while (c < 4) {
      if (c == 2) break;
      c++;
    }
    outer++;
    r++;
  }
  print(outer);
}"#,
        ["2"]
    };

    nested_for_while_break_inner_loop => {
        r#"void main() {
  var total = 0;
  for (var i = 0; i < 2; i++) {
    var j = 0;
    while (j < 5) {
      if (j == 2) break;
      total++;
      j++;
    }
  }
  print(total);
}"#,
        ["4"]
    };

    labeled_break_outer_exits_both_loops => {
        r#"void main() {
  var count = 0;
  outer:
  for (var i = 0; i < 4; i++) {
    for (var j = 0; j < 4; j++) {
      count++;
      if (j == 1) break outer;
    }
  }
  print(count);
}"#,
        ["1"]
    };

    labeled_break_outer_on_second_outer_iteration => {
        r#"void main() {
  var count = 0;
  outer:
  for (var i = 0; i < 3; i++) {
    for (var j = 0; j < 3; j++) {
      count++;
      if (i == 1 && j == 1) break outer;
    }
  }
  print(count);
}"#,
        ["5"]
    };

    labeled_continue_outer_skips_inner_rest => {
        r#"void main() {
  outer:
  for (var i = 0; i < 3; i++) {
    for (var j = 0; j < 3; j++) {
      if (j == 1) continue outer;
      print('$i$j');
    }
  }
}"#,
        ["00", "10", "20"]
    };

    labeled_continue_outer_twice_per_row => {
        r#"void main() {
  var count = 0;
  outer:
  for (var i = 0; i < 2; i++) {
    for (var j = 0; j < 4; j++) {
      count++;
      if (j == 1) continue outer;
    }
  }
  print(count);
}"#,
        ["4"]
    };

    for_break_inside_nested_if => {
        r#"void main() {
  var steps = 0;
  for (var i = 0; i < 5; i++) {
    if (i > 0) {
      if (i == 3) break;
    }
    steps++;
  }
  print(steps);
}"#,
        ["3"]
    };

    while_continue_with_multiple_conditions => {
        r#"void main() {
  var i = 0;
  var sum = 0;
  while (i < 10) {
    i++;
    if (i % 2 == 0) continue;
    if (i % 3 == 0) continue;
    sum += i;
  }
  print(sum);
}"#,
        ["16"]
    };

    for_decrementing_break_at_zero => {
        r#"void main() {
  for (var i = 5; i > 0; i--) {
    if (i == 2) break;
    print(i);
  }
}"#,
        ["5", "4", "3"]
    };

    for_in_map_entries_continue_skips_key_a => {
        r#"void main() {
  var m = {'a': 1, 'b': 2, 'c': 3};
  var sum = 0;
  for (var e in m.entries) {
    if (e.key == 'a') continue;
    sum += e.value;
  }
  print(sum);
}"#,
        ["5"]
    };

    for_in_string_continue_skips_vowels => {
        r#"void main() {
  var out = '';
  for (var ch in 'hello') {
    if (ch == 'e' || ch == 'o') continue;
    out += ch;
  }
  print(out);
}"#,
        ["hll"]
    };

    nested_break_inner_then_outer_via_label => {
        r#"void main() {
  var printed = 0;
  outer:
  for (var i = 0; i < 2; i++) {
    for (var j = 0; j < 2; j++) {
      if (j == 0) continue;
      printed++;
      break outer;
    }
  }
  print(printed);
}"#,
        ["1"]
    };

    while_break_exits_before_increment => {
        r#"void main() {
  var i = 0;
  while (i < 5) {
    if (i == 2) break;
    i++;
  }
  print(i);
}"#,
        ["2"]
    };

    for_continue_does_not_skip_increment => {
        r#"void main() {
  var last = 0;
  for (var i = 0; i < 5; i++) {
    if (i == 2) continue;
    last = i;
  }
  print(last);
}"#,
        ["4"]
    };

    do_while_break_after_single_body_run => {
        r#"void main() {
  var ran = 0;
  do {
    ran++;
    break;
  } while (ran < 3);
  print(ran);
}"#,
        ["1"]
    };

    for_in_list_break_accumulates_partial_sum => {
        r#"void main() {
  var sum = 0;
  for (var x in [2, 4, 6, 8, 10]) {
    sum += x;
    if (sum >= 10) break;
  }
  print(sum);
}"#,
        ["12"]
    };

    nested_for_break_inner_on_diagonal => {
        r#"void main() {
  var hits = 0;
  for (var i = 0; i < 3; i++) {
    for (var j = 0; j < 3; j++) {
      if (i == j) break;
      hits++;
    }
  }
  print(hits);
}"#,
        ["3"]
    };

    labeled_break_on_while_inside_for => {
        r#"void main() {
  var count = 0;
  outer:
  for (var i = 0; i < 3; i++) {
    var j = 0;
    while (j < 3) {
      count++;
      if (j == 1) break outer;
      j++;
    }
  }
  print(count);
}"#,
        ["2"]
    };

    for_break_with_empty_body_after => {
        r#"void main() {
  var i = 0;
  for (; i < 10; i++) {
    break;
  }
  print(i);
}"#,
        ["0"]
    };

    while_continue_then_break_same_iteration => {
        r#"void main() {
  var i = 0;
  var printed = 0;
  while (i < 5) {
    i++;
    if (i == 2) continue;
    if (i == 4) break;
    printed++;
  }
  print(printed);
}"#,
        ["2"]
    };

    for_in_break_on_first_match_in_strings => {
        r#"void main() {
  var found = false;
  for (var ch in 'abracadabra') {
    if (ch == 'c') {
      found = true;
      break;
    }
  }
  print(found);
}"#,
        ["true"]
    };

    nested_continue_inner_only_affects_inner => {
        r#"void main() {
  var total = 0;
  for (var i = 0; i < 2; i++) {
    for (var j = 0; j < 4; j++) {
      if (j == 2) continue;
      total++;
    }
  }
  print(total);
}"#,
        ["6"]
    };

    for_break_on_last_valid_index => {
        r#"void main() {
  var nums = [10, 20, 30];
  var picked = 0;
  for (var i = 0; i < nums.length; i++) {
    picked = nums[i];
    if (i == nums.length - 1) break;
  }
  print(picked);
}"#,
        ["30"]
    };

    do_while_continue_skips_only_second_pass => {
        r#"void main() {
  var pass = 0;
  do {
    pass++;
    if (pass == 2) continue;
    print(pass);
  } while (pass < 4);
}"#,
        ["1", "3", "4"]
    };

    for_in_continue_on_string_spaces => {
        r#"void main() {
  var letters = 0;
  for (var ch in 'a b c') {
    if (ch == ' ') continue;
    letters++;
  }
  print(letters);
}"#,
        ["3"]
    };

    labeled_continue_on_nested_while => {
        r#"void main() {
  var count = 0;
  outer:
  for (var i = 0; i < 2; i++) {
    var j = 0;
    while (j < 3) {
      j++;
      if (j == 2) continue outer;
      count++;
    }
  }
  print(count);
}"#,
        ["2"]
    };

    for_break_after_continue_same_loop => {
        r#"void main() {
  for (var i = 1; i <= 5; i++) {
    if (i % 2 == 0) continue;
    if (i == 5) break;
    print(i);
  }
}"#,
        ["1", "3"]
    };

    while_nested_for_break_inner_counts => {
        r#"void main() {
  var w = 0;
  while (w < 2) {
    for (var k = 0; k < 4; k++) {
      if (k == 2) break;
      w++;
    }
    w++;
  }
  print(w);
}"#,
        ["3"]
    };
}
