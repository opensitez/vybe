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

#[test]
fn test_php_nullsafe_and_null_coalesce_precedence() {
    let out = run_prints(
        r#"<?php
class Profile { public string $avatar = 'avatar.png'; }
class User { public ?Profile $profile = null; }
$user = new User();
echo ($user?->profile?->avatar ?? 'default.png') . '|';
$user->profile = new Profile();
echo ($user?->profile?->avatar ?? 'default.png');
"#
    );
    assert_eq!(out, vec!["default.png|avatar.png"]);
}

#[test]
fn test_php_ternary_and_match_value_precedence() {
    let out = run_prints(
        r#"<?php
$input = 1;
$status = $input ?: 'falsy' ? 'A' : 'B';
$status2 = ($input ?: 'falsy') ? 'A' : 'B';
echo $status . '|' . $status2;
"#
    );
    assert_eq!(out, vec!["1|A"]);
}

#[test]
fn test_php_ternary_chain_right_associative() {
    let out = run_prints(
        r#"<?php
$x = 1;
echo 0 ? 'x' : $x ? 'y' : 'z';
$x = 0;
echo '|';
echo (0 ? 'x' : $x ? 'y' : 'z');
"#
    );
    assert_eq!(out, vec!["y|z"]);
}

#[test]
fn test_php_ternary_nested_with_parentheses_is_right_associative() {
    let out = run_prints(
        r#"<?php
$a = 0;
echo ($a ? 'yes' : ($a ? 'inner-yes' : 'inner-no'));
echo '|';
echo ($a ? 'yes' : $a ? 'first' : 'second');
"#
    );
    assert_eq!(out, vec!["inner-no|second"]);
}

#[test]
fn test_php_coalesce_assignment_does_not_overwrite_zero() {
    let out = run_prints(
        r#"<?php
$n = 0;
$n ??= 10;
echo $n;
echo '|';
$s = "";
$s ??= "fallback";
echo $s === "" ? 'empty' : 'filled';
"#
    );
    assert_eq!(out, vec!["0|empty"]);
}

#[test]
fn test_php_match_with_computed_subject_and_boolean_arms() {
    let out = run_prints(
        r#"<?php
$left = 4;
$right = 2;
$result = match ($left > $right) {
    true => "gt",
    false => "le",
};
echo $result;
echo '|';
echo match (($left - $right) === 2) {
    true => "two",
    false => "not-two",
};
"#
    );
    assert_eq!(out, vec!["gt|two"]);
}

#[test]
fn test_php_ternary_on_truthy_falsy_edge_values() {
    let out = run_prints(
        r#"<?php
echo ('' ?: 'fallback');
echo '|';
echo ('0' ?: 'fallback');
echo '|';
echo (0 ?: 'fallback');
echo '|';
echo (false ?: 'fallback');
echo '|';
echo (' ' ?: 'fallback');
"#
    );
    assert_eq!(out, vec!["fallback|fallback|fallback|fallback| "]);
}

#[test]
fn test_php_match_with_nested_nullsafe_subject() {
    let out = run_prints(
        r#"<?php
class Node { public ?Node $next = null; public string $name = 'root'; }
$root = new Node();
$root->next = new Node();
$root->next->name = 'leaf';
$subject = $root->next?->name;
echo match ($subject) {
    null => 'missing',
    'leaf' => 'leaf',
    default => 'other',
};
echo '|';
echo match ($root->next?->next?->name) {
    null => 'no-child',
    default => 'has-child',
};
"#
    );
    assert_eq!(out, vec!["leaf|no-child"]);
}

#[test]
fn test_php_ternary_coalescing_precedence_with_parentheses() {
    let out = run_prints(
        r#"<?php
$base = null;
$value = ($base ?? 'fallback') ?: 'second';
echo $value . '|';
$value2 = $base ?? ('fallback' ?: 'second');
echo $value2;
"#
    );
    assert_eq!(out, vec!["fallback|fallback"]);
}

#[test]
fn test_php_null_coalescing_does_not_evaluate_right_when_set() {
    let out = run_prints(
        r#"<?php
$log = [];
$right = function() use (&$log) {
    $log[] = 'rhs';
    return 'rhs-value';
};
$left = 'present';
echo ($left ?? $right()) . '|';
echo implode(',', $log);
"#
    );
    assert_eq!(out, vec!["present|"]);
}

#[test]
fn test_php_match_expression_in_ternary_subject() {
    let out = run_prints(
        r#"<?php
$grade = 88;
$status = match (true) {
    $grade >= 90 => 'top',
    $grade >= 80 => 'pass',
    default => 'fail',
};
$label = $status === 'pass' ? 'ok' : 'bad';
echo $label;
"#
    );
    assert_eq!(out, vec!["ok"]);
}

#[test]
fn test_php_short_ternary_on_unset_index_with_coalesce() {
    let out = run_prints(
        r#"<?php
$data = ['x' => null];
$x = ($data['x'] ?? 'fallback') ?: 'alt';
$y = $data['y'] ?? 'fallback';
echo $x . '|' . $y;
"#
    );
    assert_eq!(out, vec!["alt|fallback"]);
}

#[test]
fn test_php_ternary_nested_with_match_subject_side_effect() {
    let out = run_prints(
        r#"<?php
$calls = 0;
$next = function() use (&$calls) { $calls++; return 10; };
$expr = $next() > 5 ? match (10) {
    10 => 'ten',
    default => 'other',
} : 'low';
echo $expr . '|' . $calls;
"#
    );
    assert_eq!(out, vec!["ten|1"]);
}

#[test]
fn test_php_match_subject_with_logical_precedence_and_falsey() {
    let out = run_prints(
        r#"<?php
$value = 0;
echo match (true) {
    ($value > 0 && $value < 5) => 'small-positive',
    ($value === 0 && $value >= 0) => 'zero',
    default => 'other',
};
"#
    );
    assert_eq!(out, vec!["zero"]);
}
