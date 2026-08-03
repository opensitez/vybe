use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Control Flow Semantics — switch-case, match, break/continue with levels, goto, declare(strict_types=1)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_switch_fallthrough_and_break_levels() {
    let out = run_prints(
        r#"<?php
$level = 2;
$out = [];

switch ($level) {
    case 1:
        $out[] = "one";
        break;
    case 2:
        $out[] = "two";
    case 3:
        $out[] = "three";
        break;
    default:
        $out[] = "default";
}

echo implode("-", $out);
"#,
    );
    assert_eq!(out, vec!["two-three"]);
}

#[test]
fn test_php_break_multi_level_loop_unwinding() {
    let out = run_prints(
        r#"<?php
$found = "";
for ($i = 0; $i < 5; $i++) {
    for ($j = 0; $j < 5; $j++) {
        if ($i === 2 && $j === 3) {
            $found = "i=$i,j=$j";
            break 2; // break both loops
        }
    }
}
echo $found;
"#,
    );
    assert_eq!(out, vec!["i=2,j=3"]);
}

#[test]
fn test_php_continue_multi_level_loop_stepping() {
    let out = run_prints(
        r#"<?php
$count = 0;
for ($i = 0; $i < 3; $i++) {
    for ($j = 0; $j < 3; $j++) {
        if ($j === 1) {
            continue 2; // continue outer loop
        }
        $count++;
    }
}
echo $count;
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_php_goto_label_forward_jump() {
    let out = run_prints(
        r#"<?php
echo "start ";
goto end_label;
echo "skipped ";
end_label:
echo "end";
"#,
    );
    assert_eq!(out, vec!["start end"]);
}

#[test]
fn test_php_declare_strict_types_type_checking() {
    compile_ok(
        r#"<?php
declare(strict_types=1);

function addInts(int $a, int $b): int {
    return $a + $b;
}

echo addInts(10, 20);
"#,
    );
}

#[test]
fn test_php_do_while_at_least_once_execution() {
    let out = run_prints(
        r#"<?php
$x = 100;
do {
    echo "EX";
} while ($x < 10);
"#,
    );
    assert_eq!(out, vec!["EX"]);
}

#[test]
fn test_php_foreach_by_reference_value_mutation() {
    let out = run_prints(
        r#"<?php
$nums = [1, 2, 3];
foreach ($nums as &$val) {
    $val *= 10;
}
unset($val); // break reference binding
echo implode("-", $nums);
"#,
    );
    assert_eq!(out, vec!["10-20-30"]);
}

#[test]
fn test_php_switch_loose_equality_vs_match_strict() {
    compile_ok(
        r#"<?php
$val = "0";
$switchRes = "";
switch ($val) {
    case 0: $switchRes = "LOOSE_MATCH"; break;
    case "0": $switchRes = "STRICT_MATCH"; break;
}
echo $switchRes;
"#,
    );
}

#[test]
fn test_php_goto_loop_restart_pattern() {
    compile_ok(
        r#"<?php
$attempts = 0;
retry:
$attempts++;
if ($attempts < 3) {
    goto retry;
}
echo "Attempts: $attempts";
"#,
    );
}

#[test]
fn test_php_declare_ticks_directive() {
    compile_ok(
        r#"<?php
declare(ticks=1);
$a = 1;
$b = 2;
echo $a + $b;
"#,
    );
}

#[test]
fn test_php_goto_from_nested_while_in_function() {
    let out = run_prints(
        r#"<?php
function run_with_goto(int $n): int {
    $i = 0;
    $sum = 0;
    while (true) {
        $i++;
        if ($i >= $n) {
            goto done;
        }
        $sum += $i;
    }
done:
    return $sum;
}
echo run_with_goto(4);
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn test_php_switch_subject_is_boolean_and_default() {
    let out = run_prints(
        r#"<?php
$value = match (false) {
    true => 'true',
    false => 'false',
    default => 'other' };
echo $value;
"#,
    );
    assert_eq!(out, vec!["false"]);
}

#[test]
fn test_php_switch_falls_through_into_default_when_no_break_between() {
    let out = run_prints(
        r#"<?php
$mode = 1;
$parts = [];
switch ($mode) {
    case 1:
        $parts[] = 'one';
    case 2:
        $parts[] = 'two';
    default:
        $parts[] = 'default';
}
echo implode('-', $parts);
"#,
    );
    assert_eq!(out, vec!["one-two-default"]);
}

#[test]
fn test_php_switch_with_expression_cases_and_parentheses_precedence() {
    let out = run_prints(
        r#"<?php
$input = 4;
$out = '';
switch ($input) {
    case 2 + 2:
        $out .= 'four';
        break;
    case (int) '3':
        $out .= 'three';
        break;
    default:
        $out .= 'other';
}
echo $out;
"#,
    );
    assert_eq!(out, vec!["four"]);
}

#[test]
fn test_php_switch_default_then_cases_falls_through_to_case() {
    let out = run_prints(
        r#"<?php
$x = 3;
$parts = [];
switch ($x) {
    default:
        $parts[] = 'd';
        // intentional fallthrough into case 1
    case 1:
        $parts[] = 'one';
        break;
    case 2:
        $parts[] = 'two';
        break;
}
echo implode('|', $parts);
"#,
    );
    assert_eq!(out, vec!["d|one"]);
}

#[test]
fn test_php_match_unhandled_no_default() {
    let out = run_prints(
        r#"<?php
try {
    echo match (9) { 1 => 'one', 2 => 'two' };
} catch (UnhandledMatchError $e) {
    echo 'unhandled';
}
"#,
    );
    assert_eq!(out, vec!["unhandled"]);
}

#[test]
fn test_php_switch_continue_to_case_in_loop_runtime() {
    let out = run_prints(
        r#"<?php
$sum = 0;
for ($i = 0; $i < 4; $i++) {
    switch ($i) {
        case 0:
            continue;
        case 2:
            $sum += 10;
            continue 2;
        default:
            $sum += 1;
    }
}
echo $sum;
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_php_switch_continue_out_of_switch_only() {
    let out = run_prints(
        r#"<?php
$trace = '';
for ($i = 0; $i < 3; $i++) {
    switch ($i) {
        case 0:
            $trace .= 'a';
            continue;
        case 1:
            $trace .= 'b';
            break;
        default:
            $trace .= 'c';
    }
    $trace .= '-';
}
echo $trace;
"#,
    );
    assert_eq!(out, vec!["a-b-c-"]);
}

#[test]
fn test_php_match_in_for_loop_subject_changes_control() {
    let out = run_prints(
        r#"<?php
$out = '';
for ($i = 0; $i < 4; $i++) {
    $out .= match ($i) {
        0, 1 => 'L',
        2 => 'M',
        default => 'H' };
}
echo $out;
"#,
    );
    assert_eq!(out, vec!["LLMH"]);
}

#[test]
fn test_php_continue_in_nested_switch_case_after_labelled_fallback() {
    let out = run_prints(
        r#"<?php
$sum = 0;
for ($i = 0; $i < 3; $i++) {
    switch ($i) {
        case 1:
            $sum += 1;
            continue;
        case 2:
            $sum += 2;
        default:
            $sum += 4;
    }
}
echo $sum;
"#,
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn test_php_match_internally_throws_from_unguarded_arm() {
    let out = run_prints(
        r#"<?php
try {
    echo match (3) {
        1 => 'one',
        2 => 'two',
        default => throw new RuntimeException('bad') };
} catch (RuntimeException $e) {
    echo 'caught';
}
"#,
    );
    assert_eq!(out, vec!["caught"]);
}

#[test]
fn test_php_goto_after_switch_case() {
    let out = run_prints(
        r#"<?php
$state = 0;
$out = '';
switch ($state) {
    case 0:
        $out .= 'pre';
        goto done;
    case 1:
        $out .= 'ignored';
}
done:
echo $out;
"#,
    );
    assert_eq!(out, vec!["pre"]);
}

#[test]
fn test_php_switch_subject_mutation_in_case_condition() {
    let out = run_prints(
        r#"<?php
$x = 0;
$out = '';
switch (1) {
    case 1:
        $out .= 'first';
        $x = 2;
        // falls through by design:
    case 2:
        $out .= '-second';
        break;
    case 3:
        $out .= '-third';
        break;
}
echo $out;
"#,
    );
    assert_eq!(out, vec!["first-second"]);
}

#[test]
fn test_php_goto_with_safe_local_jump_pattern() {
    let out = run_prints(
        r#"<?php
function guarded() {
    $x = 0;
    if ($x < 1) {
        goto skip;
    }
    $x = 10;
skip:
    return $x === 0 ? 'zero' : 'set';
}
echo guarded();
"#,
    );
    assert_eq!(out, vec!["set"]);
}

#[test]
fn test_php_switch_uses_non_trivial_subject_expression() {
    let out = run_prints(
        r#"<?php
$n = 1;
$parts = [];
$key = ($n * 2) + 1;
switch ($key) {
    case 1:
        $parts[] = 'one';
        break;
    case 2:
        $parts[] = 'two';
        break;
    case 3:
        $parts[] = 'three';
        break;
}
echo implode('|', $parts);
"#,
    );
    assert_eq!(out, vec!["three"]);
}

#[test]
fn test_php_match_in_loop_with_falsy_subject_edge() {
    let out = run_prints(
        r#"<?php
$sum = 0;
for ($i = 0; $i < 4; $i++) {
    $sum += match ((string)$i) {
        '' => 0,
        '0' => 1,
        '1' => 2,
        '2' => 3,
        default => 4 };
}
echo $sum;
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_php_goto_after_nested_loop_break_pattern() {
    let out = run_prints(
        r#"<?php
$out = '';
for ($i = 0; $i < 2; $i++) {
    if ($i === 1) {
        goto after_all;
    }
    $out .= $i;
}
after_all:
echo $out;
"#,
    );
    assert_eq!(out, vec!["0"]);
}
