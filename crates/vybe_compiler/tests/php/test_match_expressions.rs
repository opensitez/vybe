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
        ["B"]
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
}
