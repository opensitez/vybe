//! Classic for, while, do-while, for-in, break, continue, nested loops.

dart_cases! {
    for_loop_counts_zero_to_exclusive_bound => {
        r#"void main() {
  for (var i = 0; i < 3; i++) {
    print(i);
  }
}"#,
        ["0", "1", "2"]
    };

    for_loop_zero_iterations_when_condition_false => {
        r#"void main() {
  var ran = 0;
  for (var i = 0; i < 0; i++) {
    ran++;
  }
  print(ran);
}"#,
        ["0"]
    };

    for_loop_single_iteration => {
        r#"void main() {
  var count = 0;
  for (var i = 0; i < 1; i++) {
    count++;
  }
  print(count);
}"#,
        ["1"]
    };

    for_loop_decrements_counter => {
        r#"void main() {
  for (var i = 3; i > 0; i--) {
    print(i);
  }
}"#,
        ["3", "2", "1"]
    };

    for_loop_steps_by_two => {
        r#"void main() {
  for (var i = 0; i < 10; i += 2) {
    print(i);
  }
}"#,
        ["0", "2", "4", "6", "8"]
    };

    for_loop_accumulates_sum_one_to_five => {
        r#"void main() {
  var sum = 0;
  for (var i = 1; i <= 5; i++) {
    sum += i;
  }
  print(sum);
}"#,
        ["15"]
    };

    for_loop_accumulates_product_one_to_five => {
        r#"void main() {
  var product = 1;
  for (var i = 1; i <= 5; i++) {
    product *= i;
  }
  print(product);
}"#,
        ["120"]
    };

    for_loop_continue_skips_even_numbers => {
        r#"void main() {
  for (var i = 1; i <= 6; i++) {
    if (i % 2 == 0) continue;
    print(i);
  }
}"#,
        ["1", "3", "5"]
    };

    for_loop_break_stops_at_threshold => {
        r#"void main() {
  for (var i = 1; i <= 10; i++) {
    if (i > 4) break;
    print(i);
  }
}"#,
        ["1", "2", "3", "4"]
    };

    for_loop_break_before_first_print => {
        r#"void main() {
  var printed = 0;
  for (var i = 0; i < 10; i++) {
    if (i == 0) break;
    printed++;
  }
  print(printed);
}"#,
        ["0"]
    };

    for_loop_finds_max_in_sequence => {
        r#"void main() {
  var nums = [3, 9, 1, 7, 4];
  var max = nums[0];
  for (var i = 1; i < nums.length; i++) {
    if (nums[i] > max) max = nums[i];
  }
  print(max);
}"#,
        ["9"]
    };

    for_loop_finds_min_in_sequence => {
        r#"void main() {
  var nums = [3, 9, 1, 7, 4];
  var min = nums[0];
  for (var i = 1; i < nums.length; i++) {
    if (nums[i] < min) min = nums[i];
  }
  print(min);
}"#,
        ["1"]
    };

    for_loop_builds_comma_separated_string => {
        r#"void main() {
  var parts = <String>[];
  for (var i = 0; i < 3; i++) {
    parts.add('x$i');
  }
  print(parts.join(','));
}"#,
        ["x0,x1,x2"]
    };

    for_loop_counts_matching_elements => {
        r#"void main() {
  var nums = [1, 2, 2, 3, 2, 4];
  var count = 0;
  for (var i = 0; i < nums.length; i++) {
    if (nums[i] == 2) count++;
  }
  print(count);
}"#,
        ["3"]
    };

    while_loop_counts_down_to_one => {
        r#"void main() {
  var n = 3;
  while (n > 0) {
    print(n);
    n--;
  }
}"#,
        ["3", "2", "1"]
    };

    while_loop_zero_iterations_when_condition_false => {
        r#"void main() {
  var ran = 0;
  while (false) {
    ran++;
  }
  print(ran);
}"#,
        ["0"]
    };

    while_loop_accumulates_factorial => {
        r#"void main() {
  var i = 1;
  var fact = 1;
  while (i <= 5) {
    fact *= i;
    i++;
  }
  print(fact);
}"#,
        ["120"]
    };

    while_loop_with_break_exits_early => {
        r#"void main() {
  var i = 0;
  while (true) {
    if (i >= 3) break;
    print(i);
    i++;
  }
}"#,
        ["0", "1", "2"]
    };

    while_loop_continue_skips_multiples_of_three => {
        r#"void main() {
  var i = 0;
  while (i < 8) {
    i++;
    if (i % 3 == 0) continue;
    print(i);
  }
}"#,
        ["1", "2", "4", "5", "7", "8"]
    };

    while_loop_builds_repeated_characters => {
        r#"void main() {
  var s = '';
  var i = 0;
  while (i < 4) {
    s += 'z';
    i++;
  }
  print(s);
}"#,
        ["zzzz"]
    };

    do_while_runs_body_at_least_once => {
        r#"void main() {
  var i = 0;
  do {
    print(i);
    i++;
  } while (i < 1);
}"#,
        ["0"]
    };

    do_while_repeats_until_condition_false => {
        r#"void main() {
  var i = 0;
  do {
    i++;
  } while (i < 4);
  print(i);
}"#,
        ["4"]
    };

    do_while_condition_false_after_first_body => {
        r#"void main() {
  var count = 0;
  do {
    count++;
  } while (count < 1);
  print(count);
}"#,
        ["1"]
    };

    do_while_accumulates_with_post_check => {
        r#"void main() {
  var sum = 0;
  var n = 1;
  do {
    sum += n;
    n++;
  } while (n <= 3);
  print(sum);
}"#,
        ["6"]
    };

    for_in_iterates_list_elements => {
        r#"void main() {
  for (var x in [10, 20, 30]) {
    print(x);
  }
}"#,
        ["10", "20", "30"]
    };

    for_in_empty_list_runs_zero_times => {
        r#"void main() {
  var count = 0;
  for (var x in <int>[]) {
    count++;
  }
  print(count);
}"#,
        ["0"]
    };

    for_in_sums_list_elements => {
        r#"void main() {
  var sum = 0;
  for (var x in [1, 2, 3, 4]) {
    sum += x;
  }
  print(sum);
}"#,
        ["10"]
    };

    for_in_over_string_characters => {
        r#"void main() {
  var count = 0;
  for (var ch in 'abc') {
    count++;
  }
  print(count);
}"#,
        ["3"]
    };

    for_in_map_values_via_entries => {
        r#"void main() {
  var m = {'a': 1, 'b': 2};
  var sum = 0;
  for (var e in m.entries) {
    sum += e.value;
  }
  print(sum);
}"#,
        ["3"]
    };

    for_in_set_elements => {
        r#"void main() {
  var total = 0;
  for (var x in {1, 2, 3}) {
    total += x;
  }
  print(total);
}"#,
        ["6"]
    };

    nested_for_loops_count_cartesian_pairs => {
        r#"void main() {
  var count = 0;
  for (var r = 0; r < 2; r++) {
    for (var c = 0; c < 3; c++) {
      count++;
    }
  }
  print(count);
}"#,
        ["6"]
    };

    nested_for_loops_print_row_column_indices => {
        r#"void main() {
  for (var r = 0; r < 2; r++) {
    for (var c = 0; c < 2; c++) {
      print('$r$c');
    }
  }
}"#,
        ["00", "01", "10", "11"]
    };

    nested_for_break_inner_loop_only => {
        r#"void main() {
  var count = 0;
  for (var i = 0; i < 3; i++) {
    for (var j = 0; j < 3; j++) {
      if (j == 1) break;
      count++;
    }
  }
  print(count);
}"#,
        ["3"]
    };

    nested_for_continue_in_inner_loop => {
        r#"void main() {
  var count = 0;
  for (var i = 0; i < 2; i++) {
    for (var j = 0; j < 3; j++) {
      if (j == 1) continue;
      count++;
    }
  }
  print(count);
}"#,
        ["4"]
    };

    for_then_while_mixed_control_flow => {
        r#"void main() {
  var sum = 0;
  for (var i = 1; i <= 3; i++) {
    sum += i;
  }
  var j = 0;
  while (j < 2) {
    sum += 10;
    j++;
  }
  print(sum);
}"#,
        ["26"]
    };

    for_loop_with_initialization_outside => {
        r#"void main() {
  var i = 5;
  for (; i < 8; i++) {
    print(i);
  }
}"#,
        ["5", "6", "7"]
    };

    for_loop_with_missing_update_runs_forever_until_break => {
        r#"void main() {
  var count = 0;
  for (var i = 0; i < 5; ) {
    count++;
  if (count >= 3) break;
    i++;
  }
  print(count);
}"#,
        ["3"]
    };

    while_loop_sentinel_minus_one => {
        r#"void main() {
  var data = [5, 3, 8, -1, 2];
  var idx = 0;
  var sum = 0;
  while (data[idx] != -1) {
    sum += data[idx];
    idx++;
  }
  print(sum);
}"#,
        ["16"]
    };

    for_loop_reverses_into_new_list => {
        r#"void main() {
  var src = [1, 2, 3];
  var rev = <int>[];
  for (var i = src.length - 1; i >= 0; i--) {
    rev.add(src[i]);
  }
  print(rev.join('-'));
}"#,
        ["3-2-1"]
    };

    for_loop_detects_adjacent_duplicates => {
        r#"void main() {
  var nums = [1, 2, 2, 3];
  var found = false;
  for (var i = 0; i < nums.length - 1; i++) {
    if (nums[i] == nums[i + 1]) found = true;
  }
  print(found);
}"#,
        ["true"]
    };

    for_loop_running_index_with_offset => {
        r#"void main() {
  var offset = 100;
  for (var i = 0; i < 3; i++) {
    print(offset + i);
  }
}"#,
        ["100", "101", "102"]
    };

    while_loop_doubles_until_threshold => {
        r#"void main() {
  var n = 1;
  var steps = 0;
  while (n < 20) {
    n *= 2;
    steps++;
  }
  print(steps);
}"#,
        ["5"]
    };

    do_while_reads_menu_until_quit => {
        r#"void main() {
  var choices = [1, 2, 0];
  var idx = 0;
  var picks = 0;
  do {
    picks++;
    idx++;
  } while (choices[idx - 1] != 0 && idx < choices.length);
  print(picks);
}"#,
        ["3"]
    };

    for_in_with_break_on_target => {
        r#"void main() {
  var found = 0;
  for (var x in [2, 4, 6, 8, 10]) {
    if (x == 6) {
      found = x;
      break;
    }
  }
  print(found);
}"#,
        ["6"]
    };

    for_in_with_continue_filters_negatives => {
        r#"void main() {
  var sum = 0;
  for (var x in [1, -2, 3, -4, 5]) {
    if (x < 0) continue;
    sum += x;
  }
  print(sum);
}"#,
        ["9"]
    };

    nested_while_counts_grid_cells => {
        r#"void main() {
  var rows = 2;
  var cols = 4;
  var r = 0;
  var total = 0;
  while (r < rows) {
    var c = 0;
    while (c < cols) {
      total++;
      c++;
    }
    r++;
  }
  print(total);
}"#,
        ["8"]
    };

    for_loop_modulo_pattern_prints_fizz_flags => {
        r#"void main() {
  for (var i = 1; i <= 5; i++) {
    print(i % 3 == 0);
  }
}"#,
        ["false", "false", "true", "false", "false"]
    };

    for_loop_xor_parity_toggle => {
        r#"void main() {
  var parity = 0;
  for (var i = 0; i < 4; i++) {
    parity ^= 1;
    print(parity);
  }
}"#,
        ["1", "0", "1", "0"]
    };

    while_loop_peek_then_consume_pattern => {
        r#"void main() {
  var queue = [7, 8, 9];
  var head = 0;
  var sum = 0;
  while (head < queue.length) {
    sum += queue[head];
    head++;
  }
  print(sum);
}"#,
        ["24"]
    };

    for_loop_copy_array_elements => {
        r#"void main() {
  var src = [4, 5, 6];
  var dst = List<int>.filled(src.length, 0);
  for (var i = 0; i < src.length; i++) {
    dst[i] = src[i];
  }
  print(dst[1]);
}"#,
        ["5"]
    };

    for_loop_window_of_size_two => {
        r#"void main() {
  var nums = [1, 2, 3, 4];
  var windows = 0;
  for (var i = 0; i < nums.length - 1; i++) {
    windows++;
  }
  print(windows);
}"#,
        ["3"]
    };

    break_in_do_while_loop => {
        r#"void main() {
  var i = 0;
  do {
    if (i == 2) break;
    print(i);
    i++;
  } while (i < 5);
}"#,
        ["0", "1"]
    };

    continue_in_do_while_loop => {
        r#"void main() {
  var i = 0;
  do {
    i++;
    if (i == 2) continue;
    print(i);
  } while (i < 4);
}"#,
        ["1", "3", "4"]
    };

    for_loop_leading_zero_iterations_with_negative_start => {
        r#"void main() {
  var ran = 0;
  for (var i = -3; i < -3; i++) {
    ran++;
  }
  print(ran);
}"#,
        ["0"]
    };

    for_loop_negative_range_counts_up => {
        r#"void main() {
  for (var i = -2; i <= 0; i++) {
    print(i);
  }
}"#,
        ["-2", "-1", "0"]
    };
}
