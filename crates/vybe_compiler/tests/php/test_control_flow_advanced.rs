use super::helpers::run_prints;

// ── match expression ──────────────────────────────────────────

#[test]
fn match_strict_no_coercion() {
    assert_eq!(
        run_prints(
            r#"<?php
$v = '1';
echo match($v) { 1 => 'int', '1' => 'string', default => 'other' };
"#
        ),
        vec!["string"]
    );
}
#[test]
fn match_multiple_arms_comma() {
    assert_eq!(
        run_prints(
            r#"<?php
function classify(int $n): string {
    return match(true) {
        $n < 0 => 'negative',
        $n === 0 => 'zero',
        $n < 10 => 'small',
        $n < 100 => 'medium',
        default => 'large',
    };
}
echo classify(-5) . ',' . classify(0) . ',' . classify(7) . ',' . classify(50) . ',' . classify(200);
"#
        ),
        vec!["negative,zero,small,medium,large"]
    );
}
#[test]
fn match_as_expression_in_assign() {
    assert_eq!(
        run_prints(
            r#"<?php
$status = 2;
$label = match($status) { 1 => 'active', 2 => 'pending', 3 => 'closed', default => 'unknown' };
echo $label;
"#
        ),
        vec!["pending"]
    );
}
#[test]
fn match_throws_on_no_match() {
    assert_eq!(
        run_prints(
            r#"<?php
try { $r = match(99) { 1 => 'one' }; }
catch (\UnhandledMatchError $e) { echo 'unhandled'; }
"#
        ),
        vec!["unhandled"]
    );
}
#[test]
fn match_with_no_arg_bool_conditions() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = 15;
echo match(true) { $x % 15 === 0 => 'FizzBuzz', $x % 3 === 0 => 'Fizz', $x % 5 === 0 => 'Buzz', default => (string)$x };
"#
        ),
        vec!["FizzBuzz"]
    );
}

// ── Switch fall-through ───────────────────────────────────────

#[test]
fn switch_fall_through() {
    assert_eq!(
        run_prints(
            r#"<?php
$v = 2;
switch ($v) {
    case 1:
    case 2:
    case 3:
        echo 'low';
        break;
    default:
        echo 'high';
}
"#
        ),
        vec!["low"]
    );
}
#[test]
fn switch_return_from_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function day(int $n): string {
    switch ($n) {
        case 1: return 'Mon';
        case 2: return 'Tue';
        case 3: return 'Wed';
        default: return 'Other';
    }
}
echo day(2);
"#
        ),
        vec!["Tue"]
    );
}

// ── Loop control ──────────────────────────────────────────────

#[test]
fn break_with_level() {
    assert_eq!(
        run_prints(
            r#"<?php
for ($i = 0; $i < 3; $i++) {
    for ($j = 0; $j < 3; $j++) {
        if ($j === 1) break 2;
        echo $i . $j;
    }
}
"#
        ),
        vec!["00"]
    );
}
#[test]
fn continue_with_level() {
    assert_eq!(
        run_prints(
            r#"<?php
for ($i = 0; $i < 2; $i++) {
    for ($j = 0; $j < 3; $j++) {
        if ($j === 1) continue 2;
        echo $i . $j . ',';
    }
}
"#
        ),
        vec!["00,10,"]
    );
}
#[test]
fn do_while_executes_once() {
    assert_eq!(
        run_prints(
            r#"<?php
$i = 10;
do { echo $i; $i++; } while ($i < 5);
"#
        ),
        vec!["10"]
    );
}

// ── Short-circuit evaluation ──────────────────────────────────

#[test]
fn and_short_circuit() {
    assert_eq!(
        run_prints(
            r#"<?php
$called = false;
function sideEffect(bool &$flag): bool { $flag = true; return true; }
false && sideEffect($called);
echo $called ? 'called' : 'skipped';
"#
        ),
        vec!["skipped"]
    );
}
#[test]
fn or_short_circuit() {
    assert_eq!(
        run_prints(
            r#"<?php
$called = false;
function sideEffect2(bool &$flag): bool { $flag = true; return true; }
true || sideEffect2($called);
echo $called ? 'called' : 'skipped';
"#
        ),
        vec!["skipped"]
    );
}

// ── Conditional assignment patterns ──────────────────────────

#[test]
fn ternary_nested() {
    assert_eq!(
        run_prints(
            r#"<?php
$grade = function(int $s): string {
    return $s >= 90 ? 'A' : ($s >= 80 ? 'B' : ($s >= 70 ? 'C' : 'F'));
};
echo $grade(95) . $grade(82) . $grade(75) . $grade(60);
"#
        ),
        vec!["ABCF"]
    );
}
#[test]
fn null_coalescing_chained_deep() {
    assert_eq!(
        run_prints(
            r#"<?php
$config = ['db' => ['host' => 'localhost']];
echo $config['db']['port'] ?? $config['cache']['port'] ?? 3306;
"#
        ),
        vec!["3306"]
    );
}

// ── for loop variants ─────────────────────────────────────────

#[test]
fn for_multiple_init_exprs() {
    assert_eq!(
        run_prints(
            r#"<?php
for ($i = 0, $j = 10; $i < $j; $i++, $j--) {}
echo $i . ':' . $j;
"#
        ),
        vec!["5:5"]
    );
}
#[test]
fn foreach_reference() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = [1,2,3,4,5];
foreach ($a as &$v) $v *= 2;
unset($v);
echo implode(',', $a);
"#
        ),
        vec!["2,4,6,8,10"]
    );
}
#[test]
fn foreach_with_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
$map = ['a'=>1,'b'=>2,'c'=>3];
$result = [];
foreach ($map as $k => $v) $result[] = "$k=$v";
echo implode(',', $result);
"#
        ),
        vec!["a=1,b=2,c=3"]
    );
}

// ── Recursive algorithms ──────────────────────────────────────

#[test]
fn recursive_binary_search() {
    assert_eq!(
        run_prints(
            r#"<?php
function bsearch(array $a, int $target, int $lo = 0, ?int $hi = null): int {
    $hi ??= count($a) - 1;
    if ($lo > $hi) return -1;
    $mid = intdiv($lo + $hi, 2);
    return match(true) {
        $a[$mid] === $target => $mid,
        $a[$mid] < $target  => bsearch($a, $target, $mid + 1, $hi),
        default              => bsearch($a, $target, $lo, $mid - 1),
    };
}
$sorted = range(0, 20, 2);
echo bsearch($sorted, 14);
"#
        ),
        vec!["7"]
    );
}
#[test]
fn recursive_quicksort() {
    assert_eq!(
        run_prints(
            r#"<?php
function qsort(array $a): array {
    if (count($a) <= 1) return $a;
    $pivot = $a[0];
    $left  = array_filter(array_slice($a,1), fn($x) => $x <= $pivot);
    $right = array_filter(array_slice($a,1), fn($x) => $x > $pivot);
    return [...qsort(array_values($left)), $pivot, ...qsort(array_values($right))];
}
echo implode(',', qsort([3,6,8,10,1,2,1]));
"#
        ),
        vec!["1,1,2,3,6,8,10"]
    );
}
