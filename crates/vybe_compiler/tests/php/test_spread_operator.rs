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
}
