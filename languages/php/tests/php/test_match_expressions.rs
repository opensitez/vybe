//! `match` expressions, `default`, guards, and `UnhandledMatchError`.

crate::php_cases! {
    match_literal_arm_selected => {
        r#"<?php
echo match (2) { 1 => 'one', 2 => 'two', default => 'other' };
"#,
        ["two"]
    };

    match_default_arm_when_no_literal_matches => {
        r#"<?php
echo match (99) { 1 => 'one', default => 'fallback' };
"#,
        ["fallback"]
    };

    match_multiple_values_one_arm => {
        r#"<?php
echo match (404) { 200, 201 => 'ok', 404 => 'missing', default => 'x' };
"#,
        ["missing"]
    };

    match_strict_equality_string_one_not_int_one => {
        r#"<?php
echo match ('1') { 1 => 'int', '1' => 'str', default => 'other' };
"#,
        ["str"]
    };

    match_with_comma_separated_conditions => {
        r#"<?php
$code = 301;
echo match ($code) { 200, 201, 204 => 'ok', 301, 302 => 'redirect', default => '?' };
"#,
        ["redirect"]
    };

    match_arm_expression_can_call_function => {
        r#"<?php
function label(int $n): string { return "n=$n"; }
echo match (3) { 1 => label(1), 3 => label(3), default => 'z' };
"#,
        ["n=3"]
    };

    match_true_condition_arm => {
        r#"<?php
$x = 0;
echo match (true) { $x === 0 => 'zero', default => 'nonzero' };
"#,
        ["zero"]
    };

    match_nested_inside_arm => {
        r#"<?php
echo match ('inner') {
    'inner' => match (2) { 1 => 'a', 2 => 'b', default => 'z' },
    default => 'outer',
};
"#,
        ["b"]
    };

    match_assigns_to_variable => {
        r#"<?php
$grade = match (85) { 90, 100 => 'A', 80, 89 => 'B', default => 'C' };
echo $grade;
"#,
        ["C"]
    };

    match_in_return_statement => {
        r#"<?php
function sign(int $n): string {
    return match (true) {
        $n < 0 => 'neg',
        $n > 0 => 'pos',
        default => 'zero',
    };
}
echo sign(0);
"#,
        ["zero"]
    };

    match_with_enum_case_arm => {
        r#"<?php
enum Status { case On; case Off; }
echo match (Status::On) { Status::On => 'on', Status::Off => 'off' };
"#,
        ["on"]
    };

    match_backed_enum_by_value => {
        r#"<?php
enum Color: string { case Red = 'r'; case Blue = 'b'; }
echo match (Color::Blue) { Color::Red => 'red', Color::Blue => 'blue' };
"#,
        ["blue"]
    };

    match_null_arm => {
        r#"<?php
echo match (null) { null => 'nil', default => 'val' };
"#,
        ["nil"]
    };

    match_array_identity_arm => {
        r#"<?php
$empty = [];
echo match ($empty) { [] => 'empty', default => 'other' };
"#,
        ["empty"]
    };

    match_float_arm_exact => {
        r#"<?php
echo match (1.5) { 1.0 => 'one', 1.5 => 'half', default => '?' };
"#,
        ["half"]
    };

    match_in_foreach_accumulator => {
        r#"<?php
$sum = 0;
foreach ([1, 2, 3] as $n) {
    $sum += match ($n) { 1 => 10, 2 => 20, default => $n };
}
echo $sum;
"#,
        ["33"]
    };

    match_throw_in_arm_caught => {
        r#"<?php
try {
    echo match (0) { 1 => 'ok', default => throw new RuntimeException('bad') };
} catch (RuntimeException) {
    echo 'caught';
}
"#,
        ["caught"]
    };

    match_no_matching_arm_throws_unhandled_match_error => {
        r#"<?php
try {
    match (5) { 1 => 'a', 2 => 'b' };
    echo 'ok';
} catch (UnhandledMatchError) {
    echo 'unhandled';
}
"#,
        ["unhandled"]
    };

    match_single_arm_always_selected => {
        r#"<?php
echo match (rand()) { default => 'only' };
"#,
        ["only"]
    };

    match_with_side_effect_in_condition => {
        r#"<?php
$hits = 0;
echo match (true) {
    (++$hits === 1) => 'first',
    default => 'later',
};
"#,
        ["first"]
    };

    match_expression_precedence_with_arithmetic => {
        r#"<?php
echo match (1 + 2 * 2) {
    5 => 'ok',
    default => 'bad',
};
echo '|';
echo match (1 + 2 > 4) {
    true => 'gt',
    false => 'le',
};
"#,
        ["ok|gt"]
    };

    match_null_coalesce_and_fallback => {
        r#"<?php
$value = null;
echo match ($value ?? 'fallback') {
    '' => 'empty',
    'fallback' => 'coalesced',
    default => 'other',
};
"#,
        ["coalesced"]
    };

    match_guard_like_logical_operators => {
        r#"<?php
$score = 78;
echo match (true) {
    $score > 90 && $score <= 100 => 'A',
    $score > 80 && $score <= 90 => 'B',
    $score > 70 && $score <= 80 => 'C',
    default => 'D',
};
"#,
        ["C"]
    };

    match_uses_nullsafe_expression => {
        r#"<?php
class Item {
    public function tag(): ?string { return null; }
}
$item = new Item();
echo match ($item?->tag()) {
    null => 'none',
    default => 'set',
};
"#,
        ["none"]
    };

    match_multiple_guards_for_same_condition => {
        r#"<?php
$x = 12;
echo match (12) {
    0, 1 => 'tiny',
    $x > 10 => 'bigger',
    12 => 'exact',
    default => 'none',
};
"#,
        ["bigger"]
    };

    match_with_falsy_subject_and_true_subject => {
        r#"<?php
echo match ('') {
    '' => 'empty-string',
    default => 'other',
};
echo '|';
echo match (false) {
    false => 'falsey',
    default => 'other-false',
};
"#,
        ["empty-string|falsey"]
    };

    match_on_bitwise_result => {
        r#"<?php
$left = (3 & 1) === 1;
$right = (3 | 1) === 4;
echo match (true) {
    $left => 'left-true',
    $right => 'right-true',
    default => 'none',
};
echo '|';
echo match (3 & 1) {
    1 => 'and-ok',
    default => 'and-fail',
};
"#,
        ["left-true|and-ok"]
    };

    match_nested_arrays_strict_key_matching => {
        r#"<?php
$payload = ['kind' => 'event', 'id' => 0];
echo match ($payload) {
    ['kind' => 'event'] => 'evt',
    ['kind' => 'log'] => 'log',
    default => 'other',
};
echo '|';
echo match ($payload['id'] ?? null) {
    0 => 'zero',
    null => 'null-id',
    default => 'other-id',
};
"#,
        ["evt|zero"]
    };

    match_with_function_call_in_arm => {
        r#"<?php
function score(string $name): int { return strlen($name); }
echo match (score('php')) {
    2 => 'two',
    3 => 'three',
    default => 'other',
};
"#,
        ["three"]
    };

    match_unary_negation_condition => {
        r#"<?php
$value = 0;
echo match ($value) {
    !$value => 'zero-true',
    default => 'nonzero',
};
"#,
        ["zero-true"]
    };

    match_subject_side_effect_only_once => {
        r#"<?php
$calls = 0;
$next = function() use (&$calls) { $calls++; return 2; };
echo match ($next()) {
    1 => 'one',
    2 => 'two',
    3 => 'three',
    default => 'other',
};
echo '|';
echo $calls;
"#,
        ["two|1"]
    };

    match_arms_evaluate_until_first_match => {
        r#"<?php
$log = [];
$first = function() use (&$log) { $log[] = 'first'; return 10; };
$second = function() use (&$log) { $log[] = 'second'; return 20; };
$third = function() use (&$log) { $log[] = 'third'; return 30; };
echo match (20) {
    $first() => '10',
    $second() => '20',
    $third() => '30',
    default => 'none',
};
echo '|';
echo implode(',', $log);
"#,
        ["20|first,second"]
    };

    match_subject_parenthesized_arithmetic => {
        r#"<?php
echo match (((1 + 2) * 4) - 5) {
    5 => 'five',
    7 => 'seven',
    default => 'other',
};
"#,
        ["seven"]
    };

    match_subject_via_ternary_expression => {
        r#"<?php
$is_primary = true;
echo match ($is_primary ? 2 : 1) {
    1 => 'primary-false',
    2 => 'primary-true',
    default => 'fallback',
};
"#,
        ["primary-true"]
    };
}
