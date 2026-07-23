use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Expression Operators — Ternary, Null Coalescing, Match Expressions & Short Operators
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_nested_null_coalescing_fallback() {
    let out = run_prints(
        r#"<?php
$data = ["user" => ["settings" => []]];
$theme = $data["user"]["settings"]["theme"] ?? $data["default_theme"] ?? "dark";
echo $theme;
"#,
    );
    assert_eq!(out, vec!["dark"]);
}

#[test]
fn test_php_chained_ternary_with_parentheses() {
    let out = run_prints(
        r#"<?php
$score = 85;
$grade = ($score >= 90) ? "A" : (($score >= 80) ? "B" : "C");
echo $grade;
"#,
    );
    assert_eq!(out, vec!["B"]);
}

#[test]
fn test_php_match_expression_returning_array_structure() {
    let out = run_prints(
        r#"<?php
$env = "prod";
$config = match ($env) {
    "dev" => ["debug" => true, "cache" => false],
    "prod" => ["debug" => false, "cache" => true],
    default => ["debug" => true, "cache" => true],
};

echo "debug=" . ($config["debug"] ? "1" : "0") . " cache=" . ($config["cache"] ? "1" : "0");
"#,
    );
    assert_eq!(out, vec!["debug=0 cache=1"]);
}

#[test]
fn test_php_spaceship_comparator_sorting_values() {
    let out = run_prints(
        r#"<?php
echo (1 <=> 2) . " " . (2 <=> 2) . " " . (3 <=> 2);
"#,
    );
    assert_eq!(out, vec!["-1 0 1"]);
}

#[test]
fn test_php_nullsafe_operator_in_ternary() {
    compile_ok(
        r#"<?php
class Profile { public string $avatar = "avatar.png"; }
class User { public ?Profile $profile = null; }

$user = new User();
$avatar = $user?->profile ? $user->profile->avatar : "default.png";
echo $avatar;
"#,
    );
}

#[test]
fn test_php_short_ternary_coalescing_comparison() {
    compile_ok(
        r#"<?php
$input = "0"; // string "0" is falsy in PHP
$val1 = $input ?: "fallback"; // short ternary triggers fallback
$val2 = $input ?? "fallback"; // null coalescing does NOT trigger (not null)

echo "Ternary=$val1 Coalesce=$val2";
"#,
    );
}

#[test]
fn test_php_match_expression_exhaustiveness() {
    compile_ok(
        r#"<?php
$state = 1;
$res = match ($state) {
    1 => "One",
    2 => "Two",
    default => "Other",
};
echo $res;
"#,
    );
}

#[test]
fn test_php_null_coalescing_on_array_offset() {
    compile_ok(
        r#"<?php
$arr = ["a" => 10];
echo $arr["a"] ?? 0;
echo $arr["b"] ?? 0;
"#,
    );
}

#[test]
fn test_php_null_coalescing_assignment_nested_key() {
    compile_ok(
        r#"<?php
$settings = [];
$settings["cache"]["ttl"] ??= 3600;
echo $settings["cache"]["ttl"];
"#,
    );
}

#[test]
fn test_php_match_expression_multiple_conditions_comma() {
    compile_ok(
        r#"<?php
$char = "e";
$isVowel = match (strtolower($char)) {
    "a", "e", "i", "o", "u" => true,
    default => false,
};
echo $isVowel ? "VOWEL" : "CONSONANT";
"#,
    );
}
