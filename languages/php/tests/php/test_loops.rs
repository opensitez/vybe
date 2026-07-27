//! `for`, `while`, `do-while`, and `foreach` loops — runtime output.

crate::php_cases! {
    for_loop_sums_integers => {
        r#"<?php
$sum = 0;
for ($i = 1; $i <= 5; $i++) { $sum += $i; }
echo $sum;
"#,
        ["15"]
    };

    while_loop_counts_down_to_zero => {
        r#"<?php
$n = 3;
$out = '';
while ($n > 0) { $out .= $n; $n--; }
echo $out;
"#,
        ["321"]
    };

    do_while_runs_body_at_least_once => {
        r#"<?php
$n = 0;
do { $n++; } while ($n < 1);
echo $n;
"#,
        ["1"]
    };

    foreach_concatenates_list_values => {
        r#"<?php
$out = '';
foreach ([1, 2, 3] as $v) { $out .= $v; }
echo $out;
"#,
        ["123"]
    };

    foreach_key_value_builds_pairs => {
        r#"<?php
$out = [];
foreach (['a' => 1, 'b' => 2] as $k => $v) { $out[] = "$k:$v"; }
echo implode(',', $out);
"#,
        ["a:1,b:2"]
    };

    foreach_by_reference_doubles_each_element => {
        r#"<?php
$nums = [1, 2, 3];
foreach ($nums as &$n) { $n *= 2; }
unset($n);
echo implode(',', $nums);
"#,
        ["2,4,6"]
    };

    foreach_break_stops_on_target_value => {
        r#"<?php
$out = '';
foreach ([1, 2, 3, 4] as $v) {
    if ($v === 3) { break; }
    $out .= $v;
}
echo $out;
"#,
        ["12"]
    };

    foreach_continue_skips_even_values => {
        r#"<?php
$out = '';
foreach ([1, 2, 3, 4] as $v) {
    if ($v % 2 === 0) { continue; }
    $out .= $v;
}
echo $out;
"#,
        ["13"]
    };

    nested_foreach_visits_matrix_cells => {
        r#"<?php
$sum = 0;
foreach ([[1, 2], [3, 4]] as $row) {
    foreach ($row as $cell) { $sum += $cell; }
}
echo $sum;
"#,
        ["10"]
    };

    for_break_exits_inner_loop_early => {
        r#"<?php
$hits = 0;
for ($i = 0; $i < 5; $i++) {
    for ($j = 0; $j < 5; $j++) {
        $hits++;
        if ($j === 1) { break; }
    }
}
echo $hits;
"#,
        ["10"]
    };

    while_continue_skips_multiples_of_three => {
        r#"<?php
$i = 0;
$sum = 0;
while ($i < 6) {
    $i++;
    if ($i % 3 === 0) { continue; }
    $sum += $i;
}
echo $sum;
"#,
        ["12"]
    };

    foreach_over_empty_array_runs_zero_times => {
        r#"<?php
$count = 0;
foreach ([] as $_) { $count++; }
echo $count;
"#,
        ["0"]
    };

    foreach_list_destructure_in_loop => {
        r#"<?php
$sum = 0;
foreach ([[1, 2], [3, 4]] as [$a, $b]) { $sum += $a + $b; }
echo $sum;
"#,
        ["10"]
    };

    for_loop_with_step_two => {
        r#"<?php
$out = '';
for ($i = 0; $i < 6; $i += 2) { $out .= $i; }
echo $out;
"#,
        ["024"]
    };

    foreach_string_iterates_bytes_in_php_82 => {
        r#"<?php
$out = '';
foreach ('ab' as $ch) { $out .= $ch; }
echo $out;
"#,
        ["ab"]
    };

    do_while_condition_checked_after_body => {
        r#"<?php
$n = 0;
$runs = 0;
do { $runs++; $n++; } while ($n < 2);
echo $runs;
"#,
        ["2"]
    };

    foreach_modifies_copy_without_reference => {
        r#"<?php
$src = [1, 2];
foreach ($src as $v) { $v = 9; }
echo implode(',', $src);
"#,
        ["1,2"]
    };

    for_nested_triangle_count => {
        r#"<?php
$cells = 0;
for ($r = 1; $r <= 3; $r++) {
    for ($c = 1; $c <= $r; $c++) { $cells++; }
}
echo $cells;
"#,
        ["6"]
    };

    foreach_assoc_accumulates_values => {
        r#"<?php
$total = 0;
foreach (['x' => 10, 'y' => 20] as $v) { $total += $v; }
echo $total;
"#,
        ["30"]
    };

    while_reading_digits_from_string_index => {
        r#"<?php
$s = '987';
$i = 0;
$out = '';
while ($i < strlen($s)) {
    $out .= $s[$i];
    $i++;
}
echo $out;
"#,
        ["987"]
    };

    foreach_generator_via_iterator_to_array => {
        r#"<?php
function gen(): Generator { yield 2; yield 4; }
$sum = 0;
foreach (gen() as $n) { $sum += $n; }
echo $sum;
"#,
        ["6"]
    };

    for_continue_skips_fifth_iteration => {
        r#"<?php
$out = '';
for ($i = 1; $i <= 5; $i++) {
    if ($i === 5) { continue; }
    $out .= $i;
}
echo $out;
"#,
        ["1234"]
    };

    foreach_break_two_exits_outer_via_flag => {
        r#"<?php
$stop = false;
$out = '';
foreach ([1, 2] as $a) {
    foreach ([3, 4] as $b) {
        $out .= "$a$b";
        $stop = true;
        break;
    }
    if ($stop) { break; }
}
echo $out;
"#,
        ["13"]
    };

    do_while_false_condition_still_ran_once => {
        r#"<?php
$once = false;
do { $once = true; } while (false);
echo $once ? 'yes' : 'no';
"#,
        ["yes"]
    };

    foreach_object_public_properties => {
        r#"<?php
$o = new stdClass();
$o->a = 1;
$o->b = 2;
$sum = 0;
foreach ($o as $v) { $sum += $v; }
echo $sum;
"#,
        ["3"]
    };

    for_decrement_counts_down => {
        r#"<?php
$out = '';
for ($i = 3; $i >= 1; $i--) { $out .= $i; }
echo $out;
"#,
        ["321"]
    };

    while_assign_in_condition_reads_counter => {
        r#"<?php
$items = [10, 20];
$i = 0;
$sum = 0;
while (($v = $items[$i] ?? null) !== null) {
    $sum += $v;
    $i++;
}
echo $sum;
"#,
        ["30"]
    };

    foreach_with_unset_key_during_iteration => {
        r#"<?php
$a = ['k' => 1, 'drop' => 2, 'keep' => 3];
unset($a['drop']);
$keys = [];
foreach ($a as $k => $_) { $keys[] = $k; }
echo implode(',', $keys);
"#,
        ["k,keep"]
    };

    for_empty_body_still_increments => {
        r#"<?php
for ($i = 0; $i < 3; $i++) {}
echo $i;
"#,
        ["3"]
    };

    foreach_spread_copy_of_outer_array => {
        r#"<?php
$base = [1, 2];
$copy = [...$base];
foreach ($copy as &$v) { $v = 9; }
unset($v);
echo implode(',', $base);
"#,
        ["1,2"]
    };

    for_loop_multiple_statements_in_head => {
        r#"<?php
$sum = 0;
for ($i = 0, $j = 1; $i < 4; $i++, $j += 2) { $sum += $i + $j; }
echo $sum;
"#,
        ["22"]
    };

    for_loop_step_with_alternative_increment => {
        r#"<?php
$out = '';
for ($i = 0; $i < 5; ) {
    $out .= $i;
    $i = $i + 2;
}
echo $out;
"#,
        ["024"]
    };

    while_loop_with_assignment_in_condition => {
        r#"<?php
$n = 0;
$out = 0;
while (($n = $n + 1) <= 3) { $out += $n; }
echo $out;
"#,
        ["6"]
    };

    while_loop_multiple_breaks_with_1_based_counter => {
        r#"<?php
$i = 0;
$out = '';
while (true) {
    $i++;
    if ($i === 2) { continue; }
    if ($i === 5) { break; }
    $out .= $i;
}
echo $out;
"#,
        ["134"]
    };

    do_while_mutates_string => {
        r#"<?php
$s = 'a';
$i = 0;
do {
    $s .= 'x';
    $i++;
} while ($i < 2);
echo $s;
"#,
        ["axx"]
    };

    foreach_destructuring_list_with_key => {
        r#"<?php
$items = [['k' => 'a', 'v' => 1], ['k' => 'b', 'v' => 2]];
$out = '';
foreach ($items as $idx => ['k' => $k, 'v' => $v]) {
    $out .= $idx . ':' . $k . $v . ';';
}
echo $out;
"#,
        ["0:a1;1:b2;"]
    };

    foreach_key_only_iteration => {
        r#"<?php
$vals = ['x' => 9, 'y' => 8];
$out = '';
foreach ($vals as $k => $_) { $out .= $k; }
echo $out;
"#,
        ["xy"]
    };

    foreach_break_2_breaks_outer_loop => {
        r#"<?php
$out = '';
for ($i = 0; $i < 3; $i++) {
    foreach ([$i, $i + 10] as $j) {
        $out .= $i . $j;
        if ($j === $i + 10) { break 2; }
    }
    $out .= 'x';
}
echo $out;
"#,
        ["010"]
    };

    foreach_continue_2_skips_outer_loop_iteration => {
        r#"<?php
$out = '';
for ($i = 0; $i < 4; $i++) {
    foreach ([0, 1] as $j) {
        if ($j === 0) { continue 2; }
        $out .= ($i + $j);
    }
    $out .= 'x';
}
echo $out;
"#,
        ["1234"]
    };

    foreach_iterator_stops_after_take => {
        r#"<?php
function gen(): Generator {
    yield 'a' => 1;
    yield 'b' => 2;
    yield 'c' => 3;
}
$sum = '';
foreach (gen() as $k => $v) {
    if ($k === 'c') { break; }
    $sum .= $k . $v;
}
echo $sum;
"#,
        ["a1b2"]
    };

    while_loop_false_initial_condition => {
        r#"<?php
$seen = 0;
while (false) { $seen = 1; }
do {} while (false);
echo $seen;
"#,
        ["0"]
    };

    for_loop_stops_at_half_open_range => {
        r#"<?php
$out = '';
for ($i = 0; $i < 5; $i++) {
    if ($i === 3) { break; }
    $out .= $i;
}
echo $out;
"#,
        ["012"]
    };

    for_loop_continue_uses_early_increment => {
        r#"<?php
$sum = 0;
for ($i = 0; $i < 5; $i++) {
    if ($i % 2 === 0) { continue; }
    $sum += $i;
}
echo $sum;
"#,
        ["4"]
    };

    while_loop_condition_reads_updated_value => {
        r#"<?php
$i = 0;
$sum = 0;
while ($i < 4) {
    $i += 1;
    if ($i === 3) { continue; }
    $sum += $i;
}
echo $sum;
"#,
        ["7"]
    };

    foreach_with_key_filter_and_map_style => {
        r#"<?php
$vals = ['a' => 1, 'b' => 2, 'c' => 3];
$out = [];
foreach ($vals as $k => $v) {
    if ($k === 'b') { continue; }
    $out[] = $k . '=' . $v;
}
echo implode('|', $out);
"#,
        ["a=1|c=3"]
    };

    foreach_break_2_skips_to_outer_after_hit => {
        r#"<?php
$out = '';
for ($i = 0; $i < 3; $i++) {
    foreach (['x', 'y'] as $j) {
        if ($j === 'y') { break 2; }
        $out .= $i . $j . ',';
    }
    $out .= 'inner; ';
}
echo $out;
"#,
        ["0x,"]
    };

    foreach_continue_2_skips_to_outer_with_condition => {
        r#"<?php
$out = '';
for ($i = 0; $i < 3; $i++) {
    foreach ([0, 1] as $j) {
        if ($j === 0) { continue 2; }
        $out .= $i . $j;
    }
    $out .= '*';
}
echo $out;
"#,
        ["01"]
    };

    foreach_reference_iteration_does_not_change_count => {
        r#"<?php
$items = [1, 2, 3];
$count = 0;
foreach ($items as &$n) {
    $count++;
}
unset($n);
echo $count . ':' . count($items);
"#,
        ["3:3"]
    };

    do_while_break_and_continue_mix => {
        r#"<?php
$i = 0;
$sum = 0;
do {
    $i++;
    if ($i === 2) { continue; }
    if ($i === 4) { break; }
    $sum += $i;
} while ($i < 6);
echo $sum;
"#,
        ["4"]
    };

    foreach_iterator_like_function_call => {
        r#"<?php
function chars(): iterable {
    yield 'x' => 1;
    yield 'y' => 2;
}
$sum = '';
foreach (chars() as $k => $v) { $sum .= $k . $v; }
echo $sum;
"#,
        ["x1y2"]
    };

    for_loop_empty_init_implicit_zero => {
        r#"<?php
$i = -1;
$out = '';
for (;; ) {
    $i++;
    if ($i >= 3) { break; }
    $out .= $i;
}
echo $out;
"#,
        ["012"]
    };

    while_loop_updates_condition_via_post_increment => {
        r#"<?php
$x = 0;
$sum = 0;
while ($x < 5) {
    $x += 1;
    if ($x === 3) { continue; }
    $sum += $x;
}
echo $sum;
"#,
        ["12"]
    };

    foreach_associative_key_value_traversal => {
        r#"<?php
$map = ['a' => 1, 'b' => 2, 'c' => 3];
$out = '';
foreach ($map as $k => $v) {
    $out .= $k . $v;
}
echo $out;
"#,
        ["a1b2c3"]
    };

    for_loop_array_populate_and_aggregate => {
        r#"<?php
$sum = 0;
for ($i = 0; $i < 4; $i++) {
    if ($i % 2 === 0) {
        $sum += $i;
    } else {
        $sum += 1;
    }
}
echo $sum;
"#,
        ["4"]
    };

    do_while_with_break_before_condition => {
        r#"<?php
$i = 0;
$out = '';
do {
    if ($i >= 2) { break; }
    $out .= $i;
    $i++;
} while (true);
echo $out;
"#,
        ["01"]
    };

    foreach_nested_with_reference_and_copy => {
        r#"<?php
$rows = [[1, 2], [3, 4]];
$flat = [];
foreach ($rows as $row) {
    foreach ($row as $v) {
        $flat[] = $v;
    }
}
echo implode(',', $flat);
"#,
        ["1,2,3,4"]
    };

    while_loop_body_uses_current_element_then_break => {
        r#"<?php
$nums = [1, 2, 3];
$sum = 0;
$i = 0;
while ($i < count($nums)) {
    $sum += $nums[$i];
    if ($i === 1) { break; }
    $i++;
}
echo $sum;
"#,
        ["3"]
    };

    for_loop_with_multiple_counters_and_array_read => {
        r#"<?php
$left = 0;
$right = 0;
for ($i = 0, $j = 3; $i < 3 && $j >= 0; $i++, $j--) {
    $left += $i;
    $right += $j;
}
echo $left;
echo $right;
"#,
        ["36"]
    };

    for_loop_break_three_levels => {
        r#"<?php
$out = '';
for ($i = 0; $i < 2; $i++) {
    for ($j = 0; $j < 2; $j++) {
        $k = 0;
        while (true) {
            if ($i === 1 && $j === 1 && $k === 0) { break 3; }
            $out .= $i . $j . $k;
            $k++;
            if ($k >= 2) { break; }
        }
    }
}
echo $out;
"#,
        ["000001010011100101"]
    };

    nested_continue_2_skips_mid_loop_level => {
        r#"<?php
$out = '';
for ($i = 0; $i < 3; $i++) {
    for ($j = 0; $j < 3; $j++) {
        if ($j === 1) { continue 2; }
        $out .= $i . $j;
    }
    $out .= '|';
}
echo $out;
"#,
        ["0002|1012|2022|"]
    };

    foreach_with_intermediate_break_includes_last_index => {
        r#"<?php
$values = ['a' => 1, 'b' => 2, 'c' => 3];
$out = '';
$i = 0;
foreach ($values as $value) {
    $out .= $value;
    $i++;
    if ($i === 2) { break; }
}
echo $out;
"#,
        ["12"]
    };

    do_while_collects_string_chars => {
        r#"<?php
$s = 'xy';
$i = 0;
$out = '';
do {
    $out .= $s[$i];
    $i++;
} while ($i < strlen($s));
echo $out;
"#,
        ["xy"]
    };

    foreach_loop_with_list_destructure_and_sparse_index => {
        r#"<?php
$rows = [0 => [10, 20], 1 => [30]];
$out = [];
foreach ($rows as $row) {
    [$first, $second = 0] = $row;
    $out[] = $first + $second;
}
echo implode(',', $out);
"#,
        ["30,30"]
    };

    while_with_post_loop_check_false_positive => {
        r#"<?php
$n = 0;
do {
    $n++;
} while (false);
while (0) {
    $n++;
}
echo $n;
"#,
        ["1"]
    };

    foreach_loop_continues_and_skips => {
        r#"<?php
$sum = 0;
foreach ([1, 2, 3, 4] as $v) {
    if ($v % 2 === 0) {
        continue;
    }
    $sum += $v;
}
echo $sum;
"#,
        ["4"]
    };

    foreach_with_nested_loops_break2 => {
        r#"<?php
$out = '';
for ($i = 0; $i < 3; $i++) {
    foreach ([1, 2, 3] as $j) {
        if ($i === 1 && $j === 2) {
            break 2;
        }
        $out .= "$i$j";
    }
}
echo $out;
"#,
        ["010203"]
    };

    while_with_increment_condition_and_break => {
        r#"<?php
$n = 0;
while ($n < 10) {
    $n++;
    if ($n === 3) { break; }
}
echo $n;
"#,
        ["3"]
    };

    do_while_nested_break2 => {
        r#"<?php
$out = '';
$outer = 0;
do {
    $inner = 0;
    $outer++;
    do {
        $inner++;
        if ($outer === 2 && $inner === 2) {
            break 2;
        }
        $out .= $inner;
    } while ($inner < 3);
} while ($outer < 4);
echo $out;
"#,
        ["1231"]
    };

    for_loop_with_multiple_inits_and_post => {
        r#"<?php
$out = '';
for ($i = 0, $j = 0; $i < 3; $i++, $j += 2) {
    $out .= $i . ':' . $j . ';';
}
echo $out;
"#,
        ["0:0;1:2;2:4;"]
    };

    foreach_array_access_updates_source => {
        r#"<?php
$items = [1, 2, 3];
$out = 0;
foreach ($items as &$v) {
    $v *= 10;
}
foreach ($items as $v) {
    $out += $v;
}
echo $out;
"#,
        ["60"]
    };

    for_loop_with_continue_and_nested_condition => {
        r#"<?php
$out = '';
for ($i = 1; $i <= 6; $i++) {
    if ($i % 3 === 0) { continue; }
    $out .= $i;
}
echo $out;
"#,
        ["1245"]
    };

    for_loop_with_continue_increments_before_skip => {
        r#"<?php
$sum = 0;
for ($i = 0; $i < 6; $i++) {
    $sum += $i;
    if ($i % 2 === 0) {
        continue;
    }
    $sum += 10;
}
echo $sum;
"#,
        ["45"]
    };

    while_loop_with_nested_if_and_continue => {
        r#"<?php
$i = 0;
$out = '';
while ($i < 6) {
    $i++;
    if ($i === 2 || $i === 5) {
        continue;
    }
    $out .= $i;
}
echo $out;
"#,
        ["1346"]
    };

    do_while_with_internal_break => {
        r#"<?php
$i = 0;
$out = '';
do {
    $i++;
    if ($i === 4) {
        break;
    }
    $out .= $i;
} while ($i < 8);
echo $out;
"#,
        ["123"]
    };

    foreach_reference_unset_prevents_cross_iteration_leak => {
        r#"<?php
$arr = [1, 2, 3];
foreach ($arr as &$v) {
    $v += 1;
}
unset($v);
$arr[0] = 10;
$sum = 0;
foreach ($arr as $item) {
    $sum += $item;
}
echo $sum;
"#,
        ["17"]
    };

    nested_break_with_levels_in_mix => {
        r#"<?php
$out = '';
for ($i = 0; $i < 3; $i++) {
    $j = 0;
    while ($j < 3) {
        $j++;
        if ($i === 1 && $j === 2) {
            break 2;
        }
        $out .= $i . $j;
    }
}
echo $out;
"#,
        ["01020311"]
    };

    for_loop_false_initial_condition => {
        r#"<?php
$i = 0;
for (; $i > 0; $i++) {
    echo 'never';
}
echo 'done';
"#,
        ["done"]
    };

    foreach_generates_keys_and_values_with_numeric_and_string => {
        r#"<?php
$source = [0 => 'a', 'one' => 'b', 2 => 'c'];
$out = [];
foreach ($source as $k => $v) {
    $out[] = "$k:$v";
}
echo implode('|', $out);
"#,
        ["0:a|one:b|2:c"]
    };

    while_loop_with_counter_iterations => {
        r#"<?php
$i = 0;
$hits = 0;
while ($i < 3) {
    $hits += 1;
    $i++;
}
echo $hits;
"#,
        ["3"]
    };

    foreach_with_list_destructure_and_nested_skip => {
        r#"<?php
$rows = [[1, 2], [3, 4], [5, 6]];
$sum = 0;
foreach ($rows as [$a, $b]) {
    if ($a === 3) { continue; }
    $sum += $a + $b;
}
echo $sum;
"#,
        ["21"]
    };

    do_while_with_nested_for => {
        r#"<?php
$outer = 0;
$out = '';
do {
    $outer++;
    for ($i = 0; $i < 2; $i++) {
        $out .= $outer . $i;
    }
} while ($outer < 2);
echo $out;
"#,
        ["1011"]
    };

    foreach_string_like_array_iteration_order => {
        r#"<?php
$items = ['b' => 1, 'a' => 2];
$keys = [];
foreach ($items as $k => $_) {
    $keys[] = $k;
}
echo implode(',', $keys);
"#,
        ["b,a"]
    };

    for_loop_with_initialization_expression_only => {
        r#"<?php
$i = 0;
$sum = 0;
for ( ; ; $i++) {
    if ($i >= 4) { break; }
    $sum += $i;
}
echo $sum;
"#,
        ["6"]
    };

    for_loop_with_post_statement_only => {
        r#"<?php
$i = 0;
$sum = 0;
for (; $i < 3; $i++) {
    $sum += $i;
}
echo $sum;
"#,
        ["3"]
    };

    while_loop_with_assignment_condition_and_skip => {
        r#"<?php
$values = [1, 2, 3];
$i = 0;
$sum = 0;
while (($value = $values[$i] ?? null) !== null) {
    $i++;
    if ($value === 2) { continue; }
    $sum += $value;
}
echo $sum;
"#,
        ["4"]
    };

    do_while_with_post_increment_continue => {
        r#"<?php
$i = 0;
$out = '';
do {
    $i++;
    if ($i === 2) {
        continue;
    }
    if ($i === 4) {
        break;
    }
    $out .= $i;
} while ($i < 6);
echo $out;
"#,
        ["13"]
    };

    foreach_over_generator_with_keys_and_break => {
        r#"<?php
function make_pairs(): Generator {
    yield 'a' => 1;
    yield 'b' => 2;
    yield 'c' => 3;
}
$out = '';
$count = 0;
foreach (make_pairs() as $k => $v) {
    $out .= $k . $v;
    $count++;
    if ($count === 2) {
        break;
    }
}
echo $out;
"#,
        ["a1b2"]
    };

    foreach_reference_then_value_iterates_original_values => {
        r#"<?php
$nums = [1, 2, 3];
foreach ($nums as &$n) {
    $n *= 2;
}
unset($n);
$sum = 0;
foreach ($nums as $n) {
    $sum += $n;
}
echo $sum;
"#,
        ["12"]
    };

    for_loop_condition_uses_external_mutation => {
        r#"<?php
$limit = 2;
$i = 0;
$sum = 0;
while ($limit > 0) {
    for (; $i < $limit; $i++) {
        $sum += $i;
    }
    $limit--;
    $i = 0;
}
echo $sum;
"#,
        ["1"]
    };

    while_loop_with_even_only_sum => {
        r#"<?php
$i = 0;
$sum = 0;
while (true) {
    $i++;
    if ($i % 2 === 0) { continue; }
    $sum += $i;
    if ($i >= 5) { break; }
}
echo $sum;
"#,
        ["9"]
    };

    for_loop_condition_and_operator_precedence => {
        r#"<?php
$sum = 0;
for ($i = 0; $i < 3 && $i + 1 > 0; $i++) {
    $sum += $i;
}
echo $sum;
"#,
        ["3"]
    };

    while_loop_with_assignment_in_condition_and_truthiness => {
        r#"<?php
$values = [3, 2, 0, 5];
$i = 0;
$sum = '';
while (($v = $values[$i] ?? null)) {
    $sum .= $v;
    $i++;
}
echo $sum;
"#,
        ["32"]
    };

    do_while_parentheses_guard_controls_iterations => {
        r#"<?php
$i = 0;
$out = '';
do {
    $i++;
    if (($i % 2) === 0) {
        continue;
    }
    $out .= $i;
} while ($i < 6);
echo $out;
"#,
        ["135"]
    };

    foreach_loop_with_numeric_and_string_key_filter => {
        r#"<?php
$items = [0 => 'a', 1 => 'b', 's' => 'c'];
$out = '';
foreach ($items as $k => $v) {
    if (is_int($k) && $k === 1) {
        continue;
    }
    $out .= $v;
}
echo $out;
"#,
        ["ac"]
    };

    for_nested_mutation_and_condition_recheck => {
        r#"<?php
$ok = [];
$limit = 3;
for ($i = 0; $i < $limit; $i++) {
    $ok[] = $i;
    $limit--;
}
echo implode('', $ok);
"#,
        ["01"]
    };

    for_loop_with_do_while_like_mutation => {
        r#"<?php
$sum = 0;
for ($i = 0; $i < 5; $i++) {
    $sum += $i;
    if ($i >= 2) {
        $i = 4;
    }
}
echo $sum;
"#,
        ["3"]
    };

    while_loop_uses_modified_counter_in_condition => {
        r#"<?php
$i = 0;
$sum = 0;
while ($i < 4) {
    $i += 2;
    $sum += $i;
}
echo $sum;
"#,
        ["6"]
    };

    foreach_skips_missing_key_from_destructure => {
        r#"<?php
$pairs = [[ 'k' => 'a', 'v' => 1 ], [ 'k' => 'b' ]];
$out = [];
foreach ($pairs as $pair) {
    $out[] = $pair['k'] . ':' . ($pair['v'] ?? 0);
}
echo implode('|', $out);
"#,
        ["a:1|b:0"]
    };

    while_loop_mutates_counter_in_condition => {
        r#"<?php
$i = 0;
$sum = 0;
while ($i = $i + 1) {
    $sum += $i;
    if ($i >= 3) { break; }
}
echo $sum;
"#,
        ["6"]
    };

    while_loop_with_array_shift_drains_queue => {
        r#"<?php
$queue = [1, 2, 3];
$sum = 0;
while ($queue) {
    $sum += array_shift($queue);
}
echo $sum;
"#,
        ["6"]
    };

    foreach_generator_continue_level_break_level => {
        r#"<?php
function range_iterable(int $to): Generator {
    for ($i = 0; $i <= $to; $i++) {
        yield $i;
    }
}
$out = '';
foreach (range_iterable(4) as $i) {
    if ($i === 1) { continue; }
    if ($i === 4) { break; }
    $out .= $i;
}
echo $out;
"#,
        ["023"]
    };

    for_loop_break_continue_when_head_has_side_effect => {
        r#"<?php
$acc = '';
for ($i = 0, $j = [1, 2, 3]; $i < 4; $i++) {
    if ($i === 0) { $j[] = 0; continue; }
    if ($i === 3) { break; }
    $acc .= $i;
}
echo $acc;
echo '|';
echo count($j);
"#,
        ["12|4"]
    };

    do_while_with_nested_condition_after_iteration => {
        r#"<?php
$i = 0;
$out = '';
do {
    $out .= $i;
    $i++;
    if ($i % 2 === 0) {
        $out .= 'e';
    } else {
        $out .= 'o';
    }
} while ($i < 4);
echo $out;
"#,
        ["0o1o2e3o"]
    };

    while_loop_with_array_pop_empty_stop_by_false => {
        r#"<?php
$stack = [3, 2];
$out = 0;
while (($v = array_pop($stack)) !== null) {
    $out += $v;
}
echo $out;
"#,
        ["5"]
    };

    foreach_with_nested_list_destroys_reference_after_loop => {
        r#"<?php
$items = [[1], [2], [3]];
$sum = 0;
foreach ($items as &$pair) {
    $pair[0] *= 2;
}
unset($pair);
foreach ($items as $pair) {
    $sum += $pair[0];
}
echo $sum;
"#,
        ["7"]
    };

    foreach_continue_two_skips_inner_body => {
        r#"<?php
$total = '';
for ($i = 0; $i < 3; $i++) {
    foreach ([0, 1] as $j) {
        if ($j === 0) { continue 2; }
        $total .= $i . $j;
    }
    $total .= 'x';
}
echo $total;
"#,
        ["0x1x2x"]
    };

    for_loop_with_multiple_init_and_boolean_expression => {
        r#"<?php
$acc = 0;
for ($i = 0, $run = true; $run; $run = false) {
    $acc += 3;
}
echo $acc;
"#,
        ["3"]
    };

    do_while_break_after_continue_and_break => {
        r#"<?php
$i = 0;
$total = 0;
do {
    $i++;
    if ($i === 1) { continue; }
    $total += $i;
    if ($i === 4) { break; }
} while ($i < 6);
echo $total;
"#,
        ["9"]
    };

    while_condition_checks_truthiness_of_array_reference => {
        r#"<?php
$q = [1, 2];
$sum = 0;
while ($q) {
    $sum += array_shift($q);
}
echo $sum . '|' . (empty($q) ? 'empty' : 'not-empty');
"#,
        ["3|empty"]
    };

    foreach_over_range_generator_with_numeric_string_keys => {
        r#"<?php
function gen(): Generator {
    yield '0' => 'a';
    yield '1' => 'b';
}
$out = '';
foreach (gen() as $k => $v) {
    $out .= $k . '-' . $v;
}
echo $out;
"#,
        ["0-a1-b"]
    };

    for_loop_nested_break_2_skips_outer_levels => {
        r#"<?php
$out = '';
for ($i = 0; $i < 5; $i++) {
    for ($j = 0; $j < 4; $j++) {
        if ($i === 2 && $j === 1) { break 2; }
        $out .= $i . $j;
    }
}
echo $out;
"#,
        ["000102031011121320"]
    };

    while_break_2_skips_after_partial_inner_iteration => {
        r#"<?php
$sum = 0;
$i = 0;
while (true) {
    $i++;
    foreach ([1, 2] as $v) {
        if ($i === 3 && $v === 2) {
            break 2;
        }
        $sum += $v;
    }
}
echo $sum . '|' . $i;
"#,
        ["7|3"]
    };

    do_while_with_nested_if_and_stateful_continue => {
        r#"<?php
$i = 0;
$n = 0;
do {
    $i++;
    if ($i % 2 === 1) {
        continue;
    }
    $n++;
    if ($n === 2) { break; }
} while ($i < 10);
echo $i . '|' . $n;
"#,
        ["4|2"]
    };

    foreach_with_array_shift_side_effect_condition => {
        r#"<?php
$q = [1, 2, 3, 4];
$sum = 0;
while ($q) {
    $front = array_shift($q);
    if ($front === 2) { continue; }
    $sum += $front;
}
echo $sum;
"#,
        ["8"]
    };

    for_with_empty_body_and_complex_condition => {
        r#"<?php
$i = 0;
$j = 0;
for (; $i < 6 && $j < 3; $i += 2, $j++) {}
echo $i . '|' . $j;
"#,
        ["6|3"]
    };

    for_loop_with_infinite_guard_then_break => {
        r#"<?php
$sum = 0;
$i = 0;
for (;;) {
    $sum += $i;
    $i++;
    if ($i > 4) {
        break;
    }
}
echo $sum;
"#,
        ["10"]
    };

    do_while_nested_continue_break_in_nested_switch => {
        r#"<?php
$out = 0;
$i = 0;
do {
    switch ($i) {
        case 0:
            $i++;
            continue;
        case 1:
            $out += 2;
            break;
        default:
            $out += 3;
    }
    if ($i === 2) { break; }
    $i++;
} while ($i < 5);
echo $out . ':' . $i;
"#,
        ["5:2"]
    };

    foreach_with_continue_2_into_for_loop => {
        r#"<?php
$out = '';
for ($i = 0; $i < 2; $i++) {
    foreach (['a', 'b'] as $ch) {
        if ($ch === 'a') {
            continue 2;
        }
        $out .= $i . $ch;
    }
}
echo $out;
"#,
        ["0b1b"]
    };
}
