//! Argument spread `...`, array spread `...`, and unpacking in calls.

crate::php_cases! {
    array_spread_merges_lists => {
        r#"<?php
echo implode(',', [...[1, 2], 3]);
"#,
        ["1,2,3"]
    };

    array_spread_preserves_string_keys_last_wins => {
        r#"<?php
$a = ['x' => 1, ...['x' => 2, 'y' => 3]];
echo $a['x'] . ':' . $a['y'];
"#,
        ["2:3"]
    };

    argument_spread_forwards_to_function => {
        r#"<?php
function sum(int $a, int $b, int $c): int { return $a + $b + $c; }
echo sum(...[1, 2, 3]);
"#,
        ["6"]
    };

    argument_spread_with_named_still_works => {
        r#"<?php
function pair(int $a, int $b): string { return "$a,$b"; }
echo pair(...[4, 5]);
"#,
        ["4,5"]
    };

    spread_empty_array_adds_nothing => {
        r#"<?php
echo json_encode([...[1], ...[]]);
"#,
        ["[1]"]
    };

    unpack_array_into_echo_via_splat_in_user_function => {
        r#"<?php
function join_args(string ...$parts): string { return implode('-', $parts); }
echo join_args(...['a', 'b']);
"#,
        ["a-b"]
    };

    spread_in_array_after_literal_elements => {
        r#"<?php
echo implode(',', [0, ...[1, 2]]);
"#,
        ["0,1,2"]
    };

    spread_assoc_overrides_duplicate_keys => {
        r#"<?php
$m = ['a' => 1, ...['a' => 9]];
echo $m['a'];
"#,
        ["9"]
    };

    variadic_collects_remaining_arguments => {
        r#"<?php
function tail(string $head, int ...$rest): int { return count($rest); }
echo tail('h', 1, 2, 3);
"#,
        ["3"]
    };

    spread_generator_values_into_array => {
        r#"<?php
function gen(): Generator { yield 1; yield 2; }
echo implode(',', [...gen()]);
"#,
        ["1,2"]
    };

    nested_array_spread_with_string_keys => {
        r#"<?php
$left = ['base' => 'b'];
$mid = [...$left, 'mid' => 'm'];
$right = [...$mid, 'right' => 'r', 'base' => 'override'];
echo $right['base'];
echo '|';
echo $right['mid'];
echo '|';
echo $right['right'];
"#,
        ["override|m|r"]
    };

    argument_spread_with_defaulted_parameters => {
        r#"<?php
function greet(string $name, string $title = 'mr.', string $suffix = ''): string {
    return trim("$title $name $suffix");
}
echo greet(...['Doe']);
echo '|';
echo greet(...['Jane', 'Ms.', 'Jr.']);
"#,
        ["mr. Doe|Ms. Jane Jr."]
    };

    spread_in_array_with_nested_array_literals => {
        r#"<?php
$packed = [1, ...[2, 3], ...[4], 5, ...[]];
echo implode(',', $packed);
"#,
        ["1,2,3,4,5"]
    };

    spread_array_merge_with_numeric_reindexing => {
        r#"<?php
$a = [10 => 'a'];
$b = [11 => 'b'];
$c = [...$a, ...$b];
echo count($c);
echo '|';
echo isset($c[10]) ? 'has10' : 'no10';
echo '|';
echo array_key_exists(0, $c) ? 'has0' : 'no0';
echo '|';
echo array_key_exists(1, $c) ? 'has1' : 'no1';
"#,
        ["2|no10|has0|has1"]
    };

    spread_with_generator_and_array_keys => {
        r#"<?php
function values(): Generator {
    yield 1;
    yield 2;
}
$combined = [0, ...values(), ...[3, 4]];
echo implode(',', $combined);
"#,
        ["0,1,2,3,4"]
    };
}
