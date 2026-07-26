use super::helpers::{compile_ok, run_prints};

// ── Match basics (complement to existing coverage) ─────────────

#[test]
fn match_multiple_arms_per_value() {
    compile_ok(
        r#"<?php
$code = 404;
$label = match($code) {
    200, 201, 204 => 'success',
    301, 302      => 'redirect',
    400, 401, 403 => 'client error',
    404           => 'not found',
    500, 502, 503 => 'server error',
    default       => 'unknown',
};
echo $label;
"#,
    );
}

#[test]
fn match_strict_equality() {
    compile_ok(
        r#"<?php
$val = "1";
// match uses === not ==
$result = match(true) {
    $val === 1   => 'int 1',
    $val === "1" => 'string 1',
    default      => 'other',
};
echo $result;
"#,
    );
}

#[test]
fn match_no_default_throws() {
    compile_ok(
        r#"<?php
try {
    $x = 42;
    $r = match($x) {
        1 => 'one',
        2 => 'two',
    };
} catch (\UnhandledMatchError $e) {
    echo 'unhandled match';
}
"#,
    );
}

#[test]
fn match_as_expression_in_call() {
    compile_ok(
        r#"<?php
function describe(string $s): string { return "[$s]"; }
$level = 3;
echo describe(match($level) {
    1 => 'low',
    2 => 'medium',
    3 => 'high',
    default => 'unknown',
});
"#,
    );
}

#[test]
fn match_nested() {
    compile_ok(
        r#"<?php
$type = 'user';
$role = 'admin';
$result = match($type) {
    'user' => match($role) {
        'admin'  => 'admin user',
        'editor' => 'editor user',
        default  => 'regular user',
    },
    'bot' => 'bot',
    default => 'unknown',
};
echo $result;
"#,
    );
}

#[test]
fn match_in_method() {
    compile_ok(
        r#"<?php
class Router {
    public function dispatch(string $method, string $path): string {
        return match("$method $path") {
            'GET /'       => 'home',
            'GET /users'  => 'user list',
            'POST /users' => 'create user',
            default       => '404',
        };
    }
}
$r = new Router();
echo $r->dispatch('GET', '/users');
echo $r->dispatch('POST', '/users');
echo $r->dispatch('DELETE', '/anything');
"#,
    );
}

#[test]
fn match_with_function_call_arm() {
    compile_ok(
        r#"<?php
function heavy(): string { return 'computed'; }
$flag = true;
$result = match($flag) {
    true  => heavy(),
    false => 'skipped',
};
echo $result;
"#,
    );
}

#[test]
fn match_returns_complex_value() {
    compile_ok(
        r#"<?php
$key = 'users';
$config = match($key) {
    'users'  => ['table' => 'users',  'pk' => 'id'],
    'orders' => ['table' => 'orders', 'pk' => 'order_id'],
    default  => ['table' => 'unknown', 'pk' => 'id'],
};
echo $config['table'] . ':' . $config['pk'];
"#,
    );
}

#[test]
fn match_with_enum() {
    compile_ok(
        r#"<?php
enum Suit { case Hearts; case Diamonds; case Clubs; case Spades; }
$suit = Suit::Hearts;
$color = match($suit) {
    Suit::Hearts, Suit::Diamonds => 'red',
    Suit::Clubs,  Suit::Spades   => 'black',
};
echo $color;
"#,
    );
}

#[test]
fn match_with_backed_enum() {
    compile_ok(
        r#"<?php
enum Status: string { case Active = 'A'; case Inactive = 'I'; case Pending = 'P'; }
$s = Status::Active;
$label = match($s) {
    Status::Active   => 'Active',
    Status::Inactive => 'Inactive',
    Status::Pending  => 'Pending',
};
echo $label;
"#,
    );
}

#[test]
fn match_complex_condition_via_true() {
    compile_ok(
        r#"<?php
$score = 87;
$grade = match(true) {
    $score >= 90 => 'A',
    $score >= 80 => 'B',
    $score >= 70 => 'C',
    $score >= 60 => 'D',
    default      => 'F',
};
echo $grade;
"#,
    );
}

#[test]
fn match_null_value() {
    compile_ok(
        r#"<?php
$v = null;
$result = match($v) {
    null  => 'null',
    false => 'false',
    0     => 'zero',
    ''    => 'empty string',
    default => 'something',
};
echo $result;
"#,
    );
}

#[test]
fn match_in_loop() {
    compile_ok(
        r#"<?php
$words = ['hello', 'WORLD', 'PHP', 'code'];
$result = [];
foreach ($words as $w) {
    $result[] = match(true) {
        ctype_upper($w) => strtolower($w),
        ctype_lower($w) => strtoupper($w),
        default         => $w,
    };
}
echo implode(',', $result);
"#,
    );
}

#[test]
fn match_chained_transformation() {
    compile_ok(
        r#"<?php
$input = 'EUR';
$symbol = match($input) { 'USD' => '$', 'EUR' => '€', 'GBP' => '£', default => '?' };
$rate   = match($input) { 'USD' => 1.0, 'EUR' => 1.08, 'GBP' => 1.27, default => 0.0 };
echo "$symbol" . number_format($rate, 2);
"#,
    );
}

#[test]
fn match_expression_assigned() {
    compile_ok(
        r#"<?php
declare(strict_types=1);
function mapErrorCode(int $code): string {
    return match($code) {
        1 => 'Not Found',
        2 => 'Permission Denied',
        3 => 'Timeout',
        default => "Unknown error $code",
    };
}
echo mapErrorCode(1);
echo mapErrorCode(99);
"#,
    );
}

#[test]
fn match_subject_with_computation_precedence() {
    compile_ok(
        r#"<?php
$n = 3;
$v = match (1 + 2 * $n) {
    5 => 'ok',
    7 => 'oops',
    default => 'other',
};
echo $v;
"#,
    );
}

#[test]
fn match_with_guarding_logical_conditions() {
    compile_ok(
        r#"<?php
$score = 92;
$v = match (true) {
    $score >= 90 && $score < 100 => 'high',
    $score >= 100 => 'perfect',
    default => 'low',
};
echo $v;
"#,
    );
}

#[test]
fn match_on_array_element_and_missing_key_fallback() {
    compile_ok(
        r#"<?php
$payload = ['status' => 'ok'];
$label = match ($payload['status'] ?? 'unknown') {
    'ok' => 'good',
    'fail' => 'bad',
    default => 'unknown',
};
echo $label;
echo '|';
$label2 = match ($payload['retry'] ?? null) {
    null => 'no-retry',
    0 => 'zero',
    default => 'has',
};
echo $label2;
"#,
    );
}

#[test]
fn match_with_exhaustive_boolean_chain() {
    compile_ok(
        r#"<?php
$v = match (true) {
    true === true => 'true-branch',
    false => 'false-branch',
};
echo $v;
"#,
    );
}

#[test]
fn match_only_executes_selected_arm() {
    assert_eq!(
        run_prints(
            r#"<?php
function side_effect(string &$log, string $label): string {
    $log .= $label;
    return $label;
}
$log = '';
$value = 2;
$out = match ($value) {
    1 => side_effect($log, 'A'),
    2 => 'selected',
    default => side_effect($log, 'Z'),
};
echo $out;
echo '|';
echo $log;
"#,
        ),
        vec!["selected|"]
    );
}

#[test]
fn match_stops_after_first_matching_condition_in_true_subject() {
    assert_eq!(
        run_prints(
            r#"<?php
$value = 10;
$log = '';
function mark(array &$log, string $value): string {
    $log[] = $value;
    return $value;
}
$label = match (true) {
    $value > 5 && mark($log, 'high') => 'high',
    $value > 0 && mark($log, 'positive') => 'positive',
    default => 'zero',
};
echo $label;
echo '|';
echo implode(',', $log);
"#,
        ),
        vec!["high|high"]
    );
}

#[test]
fn match_uses_fallthrough_facts_of_ternary_subject_precedence() {
    assert_eq!(
        run_prints(
            r#"<?php
$status = 3;
$label = match (1 + 1 === 2 ? $status : 0) {
    1 => 'one',
    3 => 'three',
    5 => 'five',
    default => 'other',
};
echo $label;
"#,
        ),
        vec!["three"]
    );
}

#[test]
fn match_subject_is_evaluated_once_with_side_effect() {
    assert_eq!(
        run_prints(
            r#"<?php
$count = 0;
$next = function() use (&$count) {
    $count++;
    return 2;
};
$result = match ($next()) {
    1 => 'one',
    2 => 'two',
    default => 'other',
};
echo $result;
echo '|';
echo $count;
"#,
        ),
        vec!["two|1"]
    );
}

#[test]
fn match_all_arms_with_logical_operators_and_short_circuit() {
    assert_eq!(
        run_prints(
            r#"<?php
$calls = 0;
$mark = function() use (&$calls) { $calls++; return true; };
$label = match (true) {
    false && $mark() => 'never',
    true && $mark()  => 'hit',
    default => 'none',
};
echo $label;
echo '|';
echo $calls;
"#,
        ),
        vec!["hit|1"]
    );
}

#[test]
fn match_expression_subject_with_nested_ternary() {
    assert_eq!(
        run_prints(
            r#"<?php
$n = 4;
$label = match (($n > 3) ? 'big' : 'small') {
    'small' => 'S',
    'big' => 'B',
    default => 'U',
};
echo $label;
"#,
        ),
        vec!["B"]
    );
}

#[test]
fn match_default_after_earlier_default_not_taken_when_match_found() {
    assert_eq!(
        run_prints(
            r#"<?php
$result = match (7) {
    1, 2, 3 => 'low',
    7, 8 => 'mid',
    default => 'fallback',
};
echo $result;
"#,
        ),
        vec!["mid"]
    );
}

#[test]
fn match_with_subject_on_expression_with_error_value() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = ['v' => 0];
$label = match ($x['v'] ?? null) {
    1 => 'one',
    null => 'missing',
    default => 'other',
};
echo $label;
"#,
        ),
        vec!["other"]
    );
}
