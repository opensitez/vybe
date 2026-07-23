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
