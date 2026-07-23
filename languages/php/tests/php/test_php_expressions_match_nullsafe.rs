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
