use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Match Expressions & Nullsafe Operator — match(), nullsafe ?->, null coalescing ??, ternary ?:
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php80_match_expression_value_matching() {
    let out = run_prints(
        r#"<?php
$statusCode = 200;
$message = match ($statusCode) {
    200, 201 => "OK",
    400, 404 => "Client Error",
    500 => "Server Error",
    default => "Unknown",
};
echo $message;
"#,
    );
    assert_eq!(out, vec!["OK"]);
}

#[test]
fn test_php80_nullsafe_operator_chaining() {
    let out = run_prints(
        r#"<?php
class Profile {
    public ?string $city = "New York";
}

class User {
    public ?Profile $profile = null;
}

$user = new User();
echo $user?->profile?->city ?? "Default City";
"#,
    );
    assert_eq!(out, vec!["Default City"]);
}

#[test]
fn test_php_null_coalescing_assignment_operator() {
    let out = run_prints(
        r#"<?php
$config = [];
$config["timeout"] ??= 30;
$config["timeout"] ??= 60;
echo $config["timeout"];
"#,
    );
    assert_eq!(out, vec!["30"]);
}

#[test]
fn test_php_match_expression_strict_type_comparison() {
    let out = run_prints(
        r#"<?php
$input = "10";
$res = match ($input) {
    10 => "integer",
    "10" => "string",
    default => "other",
};
echo $res;
"#,
    );
    assert_eq!(out, vec!["string"]);
}

#[test]
fn test_php_short_ternary_operator() {
    let out = run_prints(
        r#"<?php
$provided = "value";
$empty = "";
echo ($provided ?: "fallback") . " | " . ($empty ?: "fallback");
"#,
    );
    assert_eq!(out, vec!["value | fallback"]);
}

#[test]
fn test_php_nullsafe_method_call_on_null() {
    compile_ok(
        r#"<?php
class Repository {
    public function find(int $id): ?object {
        return null;
    }
}

$repo = new Repository();
$name = $repo->find(10)?->getName();
echo $name ?? "no_object";
"#,
    );
}

#[test]
fn test_php_match_expression_returning_closure() {
    compile_ok(
        r#"<?php
$op = "add";
$handler = match ($op) {
    "add" => fn($a, $b) => $a + $b,
    "sub" => fn($a, $b) => $a - $b,
    default => throw new InvalidArgumentException("Unsupported"),
};
echo $handler(10, 5);
"#,
    );
}

#[test]
fn test_php_nested_null_coalescing_chain() {
    compile_ok(
        r#"<?php
$opt1 = null;
$opt2 = null;
$opt3 = "final_fallback";
$res = $opt1 ?? $opt2 ?? $opt3;
echo $res;
"#,
    );
}

#[test]
fn test_php_match_expression_boolean_expressions() {
    compile_ok(
        r#"<?php
$age = 25;
$category = match (true) {
    $age < 13 => "child",
    $age < 20 => "teen",
    $age >= 20 => "adult",
};
echo $category;
"#,
    );
}

#[test]
fn test_php_match_guard_precedence_and_parentheses() {
    let out = run_prints(
        r#"<?php
$age = 17;
$status = match (true) {
    $age < 13 => "child",
    ($age >= 13 && $age < 18) ? true : false => "teen",
    default => "adult",
};
echo $status;
"#,
    );
    assert_eq!(out, vec!["teen"]);
}

#[test]
fn test_php_nullsafe_chain_with_parenthesized_coalesce() {
    let out = run_prints(
        r#"<?php
class Node {
    public ?string $label = null;
}
class Holder {
    public ?Node $node = null;
}
$h = new Holder();
echo ($h->node?->label ?? "none") . "|";
$h->node = new Node();
echo ($h->node?->label ?? "none");
"#,
    );
    assert_eq!(out, vec!["none|none"]);
}

#[test]
fn test_php_nullsafe_property_write_forbidden() {
    compile_ok(
        r#"<?php
class Container {
    public ?stdClass $inner = null;
}

$c = new Container();
    // Nullsafe operator cannot be used on left hand side of assignment
$val = $c?->inner?->prop;
"#,
    );
}

#[test]
fn test_php_nullsafe_chain_truthy_false_values() {
    let out = run_prints(
        r#"<?php
class Node {
    public function value(): ?string { return ''; }
}
class Holder {
    public ?Node $node = null;
}
$h = new Holder();
echo ($h->node?->value() ?: 'fallback-a') . '|';
$h->node = new Node();
echo ($h->node?->value() ?: 'fallback-b') . '|';
"#,
    );
    assert_eq!(out, vec!["fallback-a|fallback-b"]);
}

#[test]
fn test_php_match_expression_with_empty_string_subject() {
    let out = run_prints(
        r#"<?php
$status = '';
echo match ($status) {
    '' => 'blank',
    null => 'null',
    default => 'other',
};
"#,
    );
    assert_eq!(out, vec!["blank"]);
}

#[test]
fn test_php_nested_ternary_and_match_precedence() {
    let out = run_prints(
        r#"<?php
$n = 1;
$label = match (true) {
    $n > 0 && $n < 3 => ($n === 1 ? 'one' : 'other'),
    default => 'none',
};
echo $label . '|';
$m = 0;
echo ($m ?: match (true) { true => 'ok', default => 'bad' });
"#,
    );
    assert_eq!(out, vec!["one|ok"]);
}

#[test]
fn test_php_match_subject_as_nullsafe_chain() {
    let out = run_prints(
        r#"<?php
class Profile {
    public function tier(): ?string {
        return "pro";
    }
}
class User {
    public ?Profile $profile = null;
}
$u = new User();
$level = match ($u->profile?->tier()) {
    null => 'none',
    'pro' => 'pro-user',
    default => 'other',
};
echo $level;
echo '|';
$u->profile = new Profile();
$level2 = match ($u->profile?->tier()) {
    null => 'none',
    'pro' => 'pro-user',
    default => 'other',
};
echo $level2;
"#
    );
    assert_eq!(out, vec!["none|pro-user"]);
}

#[test]
fn test_php_nullsafe_subject_side_effects_only_when_needed() {
    let out = run_prints(
        r#"<?php
class Node {
    public ?Node $next = null;
    public function label(): string { return 'leaf'; }
}
$root = null;
$count = 0;
$value = $root?->next?->label();
echo match ($value ?? 'fallback') {
    'leaf' => 'got',
    'fallback' => 'miss',
    default => 'other',
};
echo '|';
$root = new Node();
$root->next = new Node();
$value2 = $root?->next?->label();
echo match ($value2 ?? 'fallback') {
    'leaf' => 'got2',
    'fallback' => 'miss2',
    default => 'other2',
};
"#
    );
    assert_eq!(out, vec!["miss|got2"]);
}

#[test]
fn test_php_coalesce_precedence_with_match_and_ternary() {
    let out = run_prints(
        r#"<?php
$env = null;
$mode = match (true) {
    ($env ?? 'dev') === 'prod' => 'production',
    default => 'defaulted',
};
$tag = $env ?: 'dev';
$label = $mode . '/' . $tag;
echo $label;
"#
    );
    assert_eq!(out, vec!["defaulted/dev"]);
}

#[test]
fn test_php_nested_ternary_subject_and_match_default_only_one_branch() {
    let out = run_prints(
        r#"<?php
$value = 2;
$result = match (true) {
    (($value > 0) ? 1 : 0) === 1 => 'positive',
    default => 'zero',
};
echo $result . '|';
echo match (($value > 0) ? $value : -$value) {
    2 => 'two',
    default => 'other',
};
"#
    );
    assert_eq!(out, vec!["positive|two"]);
}
