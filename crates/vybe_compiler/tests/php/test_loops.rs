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
        ["5"]
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
}
