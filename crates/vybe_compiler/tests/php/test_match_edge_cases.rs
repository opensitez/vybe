use super::helpers::run_prints;

// ── UnhandledMatchError when no arm matches ───────────────────

#[test]
fn match_throws_unhandled_match_error_on_no_arm() {
    assert_eq!(
        run_prints(
            r#"<?php
try {
    $x = 99;
    $result = match($x) { 1 => 'one', 2 => 'two' };
} catch (\UnhandledMatchError $e) {
    echo "unhandled";
}
"#
        ),
        vec!["unhandled"]
    );
}

#[test]
fn match_no_default_throws_for_null() {
    assert_eq!(
        run_prints(
            r#"<?php
try {
    match(null) { 0 => 'zero', false => 'false' };
} catch (\UnhandledMatchError $e) {
    echo "no match for null";
}
"#
        ),
        vec!["no match for null"]
    );
}

// ── Multiple conditions per arm ───────────────────────────────

#[test]
fn match_multiple_conditions_on_one_arm() {
    assert_eq!(
        run_prints(
            r#"<?php
function classify(int $n): string {
    return match(true) {
        $n < 0   => 'negative',
        $n === 0 => 'zero',
        $n < 10  => 'small',
        $n < 100 => 'medium',
        default  => 'large',
    };
}
echo classify(-5) . ',' . classify(0) . ',' . classify(7) . ',' . classify(50) . ',' . classify(200);
"#
        ),
        vec!["negative,zero,small,medium,large"]
    );
}

#[test]
fn match_comma_separated_values_single_arm() {
    assert_eq!(
        run_prints(
            r#"<?php
function dayType(string $day): string {
    return match($day) {
        'Saturday', 'Sunday' => 'weekend',
        'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday' => 'weekday',
        default => 'unknown',
    };
}
echo dayType('Saturday') . ',' . dayType('Monday');
"#
        ),
        vec!["weekend,weekday"]
    );
}

// ── match uses strict comparison ──────────────────────────────

#[test]
fn match_strict_comparison_does_not_coerce() {
    assert_eq!(
        run_prints(
            r#"<?php
$val = "1";
echo match($val) {
    1 => 'int one',
    "1" => 'string one',
    default => 'other',
};
"#
        ),
        vec!["string one"]
    );
}

#[test]
fn match_strict_null_not_equal_to_zero() {
    assert_eq!(
        run_prints(
            r#"<?php
echo match(null) {
    0 => 'zero',
    null => 'null',
    default => 'other',
};
"#
        ),
        vec!["null"]
    );
}

#[test]
fn match_strict_false_not_equal_to_empty_string() {
    assert_eq!(
        run_prints(
            r#"<?php
echo match(false) {
    '' => 'empty',
    0 => 'zero',
    false => 'false',
    default => 'other',
};
"#
        ),
        vec!["false"]
    );
}

// ── match as expression ───────────────────────────────────────

#[test]
fn match_as_expression_in_assignment() {
    assert_eq!(
        run_prints(
            r#"<?php
$code = 404;
$msg = match($code) { 200 => 'OK', 404 => 'Not Found', 500 => 'Error', default => 'Unknown' };
echo $msg;
"#
        ),
        vec!["Not Found"]
    );
}

#[test]
fn match_as_expression_in_return() {
    assert_eq!(
        run_prints(
            r#"<?php
function httpText(int $code): string {
    return match($code) { 200 => 'OK', 201 => 'Created', 204 => 'No Content', default => 'Unknown' };
}
echo httpText(201);
"#
        ),
        vec!["Created"]
    );
}

#[test]
fn match_as_expression_in_echo() {
    assert_eq!(
        run_prints(
            r#"<?php
$day = 3;
echo match($day) { 1 => 'Mon', 2 => 'Tue', 3 => 'Wed', 4 => 'Thu', 5 => 'Fri', default => 'Weekend' };
"#
        ),
        vec!["Wed"]
    );
}

#[test]
fn match_as_function_argument() {
    assert_eq!(
        run_prints(
            r#"<?php
$status = 'active';
echo strtoupper(match($status) { 'active' => 'running', 'stopped' => 'halted', default => 'unknown' });
"#
        ),
        vec!["RUNNING"]
    );
}

// ── Nested match ──────────────────────────────────────────────

#[test]
fn nested_match_expressions() {
    assert_eq!(
        run_prints(
            r#"<?php
$type = 'http';
$code = 200;
echo match($type) {
    'http' => match($code) { 200 => 'OK', 404 => 'Not Found', default => 'Other HTTP' },
    'ftp' => 'FTP',
    default => 'Unknown',
};
"#
        ),
        vec!["OK"]
    );
}

// ── match with complex subject ────────────────────────────────

#[test]
fn match_on_function_call_result() {
    assert_eq!(
        run_prints(
            r#"<?php
echo match(strlen("hello")) { 3 => 'short', 5 => 'medium', default => 'other' };
"#
        ),
        vec!["medium"]
    );
}

#[test]
fn match_on_ternary_result() {
    assert_eq!(
        run_prints(
            r#"<?php
$n = 7;
echo match($n > 5 ? 'big' : 'small') { 'big' => 'large number', 'small' => 'tiny number' };
"#
        ),
        vec!["large number"]
    );
}

// ── match with enum ───────────────────────────────────────────

#[test]
fn match_with_enum_arms() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Status { case Active; case Pending; case Closed; }
$s = Status::Pending;
echo match($s) {
    Status::Active => 'running',
    Status::Pending => 'waiting',
    Status::Closed => 'done',
};
"#
        ),
        vec!["waiting"]
    );
}

// ── match default arm ─────────────────────────────────────────

#[test]
fn match_default_arm_catches_all_unmatched() {
    assert_eq!(
        run_prints(
            r#"<?php
for ($i = 1; $i <= 3; $i++) {
    echo match($i) { 1 => 'one', default => 'many' } . ',';
}
"#
        ),
        vec!["one,many,many,"]
    );
}

// ── match with no-op arm (null result) ────────────────────────

#[test]
fn match_with_null_arm_result() {
    assert_eq!(
        run_prints(
            r#"<?php
$v = 2;
$result = match($v) { 1 => 'one', 2 => null, default => 'other' };
echo var_export($result, true);
"#
        ),
        vec!["NULL"]
    );
}

// ── match in loop ─────────────────────────────────────────────

#[test]
fn match_called_in_foreach() {
    assert_eq!(
        run_prints(
            r#"<?php
$grades = ['A', 'B', 'C', 'F'];
$labels = array_map(fn($g) => match($g) {
    'A' => 'excellent',
    'B' => 'good',
    'C' => 'average',
    default => 'fail',
}, $grades);
echo implode(',', $labels);
"#
        ),
        vec!["excellent,good,average,fail"]
    );
}

// ── match vs switch type comparison ──────────────────────────

#[test]
fn match_differs_from_switch_on_string_int_coercion() {
    assert_eq!(
        run_prints(
            r#"<?php
$val = 0;
echo match($val) {
    false => 'false',
    null => 'null',
    '' => 'empty',
    0 => 'zero',
    default => 'other',
};
"#
        ),
        vec!["zero"]
    );
}

// ── match with boolean subject ────────────────────────────────

#[test]
fn match_true_subject_for_conditions() {
    assert_eq!(
        run_prints(
            r#"<?php
$score = 75;
echo match(true) {
    $score >= 90 => 'A',
    $score >= 80 => 'B',
    $score >= 70 => 'C',
    $score >= 60 => 'D',
    default => 'F',
};
"#
        ),
        vec!["C"]
    );
}

// ── match with thrown exception in arm ───────────────────────

#[test]
fn match_arm_can_throw_exception() {
    assert_eq!(
        run_prints(
            r#"<?php
function requireStatus(string $s): string {
    return match($s) {
        'active' => 'ok',
        default => throw new \InvalidArgumentException("bad: $s"),
    };
}
try {
    requireStatus('unknown');
} catch (\InvalidArgumentException $e) {
    echo $e->getMessage();
}
"#
        ),
        vec!["bad: unknown"]
    );
}

// ── match arm side effects ────────────────────────────────────

#[test]
fn match_evaluates_only_matching_arm() {
    assert_eq!(
        run_prints(
            r#"<?php
$calls = 0;
$increment = function() use (&$calls) { $calls++; return 'called'; };
$result = match(2) {
    1 => $increment(),
    2 => 'two',
    3 => $increment(),
};
echo "$result,$calls";
"#
        ),
        vec!["two,0"]
    );
}
