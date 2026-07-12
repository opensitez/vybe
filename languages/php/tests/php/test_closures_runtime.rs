//! Closure `use`, arrow functions, and callable patterns — not `Closure::bind` (see `test_callables.rs`).

crate::php_cases! {
    closure_use_imports_outer_variable_by_value => {
        r#"<?php
$base = 10;
$add = function (int $n) use ($base): int { return $base + $n; };
echo $add(5);
"#,
        ["15"]
    };

    closure_use_by_reference_mutates_outer => {
        r#"<?php
$total = 0;
$inc = function (int $n) use (&$total): void { $total += $n; };
$inc(3);
$inc(4);
echo $total;
"#,
        ["7"]
    };

    closure_static_variable_persists_between_calls => {
        r#"<?php
$tick = function (): int {
    static $n = 0;
    return ++$n;
};
echo $tick() . $tick() . $tick();
"#,
        ["123"]
    };

    arrow_function_captures_local_variable => {
        r#"<?php
$factor = 3;
$mul = fn(int $x) => $x * $factor;
echo $mul(4);
"#,
        ["12"]
    };

    arrow_function_single_expression_no_use_needed_for_globals => {
        r#"<?php
$len = fn(string $s) => strlen($s);
echo $len('php');
"#,
        ["3"]
    };

    closure_passed_to_array_map => {
        r#"<?php
echo implode(',', array_map(fn($x) => $x * 2, [1, 2, 3]));
"#,
        ["2,4,6"]
    };

    closure_passed_to_array_filter => {
        r#"<?php
echo implode(',', array_filter([1, 2, 3, 4], fn($n) => $n % 2 === 0));
"#,
        ["2,4"]
    };

    closure_passed_to_array_reduce => {
        r#"<?php
echo array_reduce([1, 2, 3], fn($c, $n) => $c + $n, 0);
"#,
        ["6"]
    };

    closure_returns_another_closure => {
        r#"<?php
$maker = function (int $offset): callable {
    return fn(int $x) => $x + $offset;
};
echo $maker(5)(10);
"#,
        ["15"]
    };

    call_user_func_invokes_closure => {
        r#"<?php
echo call_user_func(fn($a, $b) => $a . $b, 'x', 'y');
"#,
        ["xy"]
    };

    call_user_func_array_spreads_arguments => {
        r#"<?php
echo call_user_func_array(fn($a, $b) => $a - $b, [9, 4]);
"#,
        ["5"]
    };

    closure_inside_foreach_captures_loop_index => {
        r#"<?php
$out = [];
foreach ([10, 20] as $i => $v) {
    $out[] = (function () use ($i, $v) { return $i . ':' . $v; })();
}
echo implode(',', $out);
"#,
        ["0:10,1:20"]
    };

    closure_typed_parameters_and_return => {
        r#"<?php
$parse = function (string $s): int { return (int)$s; };
echo $parse('42');
"#,
        ["42"]
    };

    closure_as_object_invoke_via_callable_syntax => {
        r#"<?php
$inv = new class {
    public function __invoke(int $n): int { return $n + 1; }
};
echo $inv(4);
"#,
        ["5"]
    };

    array_walk_with_closure_mutates_array => {
        r#"<?php
$a = [1, 2, 3];
array_walk($a, function (&$v) { $v *= 10; });
echo implode(',', $a);
"#,
        ["10,20,30"]
    };

    usort_with_closure_orders_descending => {
        r#"<?php
$a = [3, 1, 2];
usort($a, fn($x, $y) => $y <=> $x);
echo implode(',', $a);
"#,
        ["3,2,1"]
    };

    closure_recursive_via_reference_use => {
        r#"<?php
$fact = null;
$fact = function (int $n) use (&$fact): int {
    return $n <= 1 ? 1 : $n * $fact($n - 1);
};
echo $fact(5);
"#,
        ["120"]
    };

    closure_match_expression_arm => {
        r#"<?php
$label = (fn(string $c) => match ($c) { 'a' => 'alpha', default => 'other' })('a');
echo $label;
"#,
        ["alpha"]
    };

    closure_in_array_stored_and_called_later => {
        r#"<?php
$ops = ['dbl' => fn($x) => $x * 2];
echo $ops['dbl'](6);
"#,
        ["12"]
    };

    closure_default_parameter_value => {
        r#"<?php
$greet = function (string $who = 'world'): string { return "hi:$who"; };
echo $greet();
"#,
        ["hi:world"]
    };

    closure_variadic_collects_arguments => {
        r#"<?php
$sum = function (int ...$nums): int { return array_sum($nums); };
echo $sum(1, 2, 3);
"#,
        ["6"]
    };

    first_class_callable_from_closure => {
        r#"<?php
$fn = strlen(...);
echo $fn('abc');
"#,
        ["3"]
    };

    closure_nullable_return_type => {
        r#"<?php
$maybe = function (bool $ok): ?string { return $ok ? 'yes' : null; };
echo $maybe(false) === null ? 'null' : 'val';
"#,
        ["null"]
    };

    closure_union_parameter_type => {
        r#"<?php
$show = function (int|string $v): string { return (string)$v; };
echo $show(7);
"#,
        ["7"]
    };

    generator_closure_yields_from_foreach => {
        r#"<?php
$gen = function (): Generator {
    foreach ([1, 2] as $n) { yield $n; }
};
echo implode('', iterator_to_array($gen()));
"#,
        ["12"]
    };

    closure_bound_to_global_function_scope_still_runs => {
        r#"<?php
$run = function (): string { return 'ok'; };
echo $run();
"#,
        ["ok"]
    };

    array_map_with_two_arrays_closure => {
        r#"<?php
echo implode(',', array_map(fn($a, $b) => $a + $b, [1, 2], [10, 20]));
"#,
        ["11,22"]
    };

    closure_early_return_inside => {
        r#"<?php
$pick = function (int $n): string {
    if ($n < 0) return 'neg';
    return 'pos';
};
echo $pick(1);
"#,
        ["pos"]
    };

    closure_captures_object_property_read => {
        r#"<?php
class Bag { public function __construct(public int $n) {} }
$b = new Bag(8);
$read = function () use ($b): int { return $b->n; };
echo $read();
"#,
        ["8"]
    };

    array_reduce_builds_comma_string => {
        r#"<?php
echo array_reduce(['a', 'b', 'c'], fn($c, $i) => $c === '' ? $i : "$c,$i", '');
"#,
        ["a,b,c"]
    };

    closure_passed_to_register_shutdown_not_run_yet => {
        r#"<?php
$ran = 'no';
$fn = function () use (&$ran) { $ran = 'yes'; };
$fn();
echo $ran;
"#,
        ["yes"]
    };

    fn_arrow_cannot_use_yield_so_returns_value => {
        r#"<?php
$double = fn($x) => $x * 2;
echo $double(3);
"#,
        ["6"]
    };

    closure_spread_operator_in_call => {
        r#"<?php
$add = function (int $a, int $b, int $c): int { return $a + $b + $c; };
echo $add(...[1, 2, 3]);
"#,
        ["6"]
    };

    array_filter_preserves_keys_with_closure => {
        r#"<?php
$a = ['x' => 1, 'y' => 0, 'z' => 3];
echo implode(',', array_keys(array_filter($a, fn($v) => $v > 0)));
"#,
        ["x,z"]
    };

    closure_composes_two_functions => {
        r#"<?php
$compose = function (callable $f, callable $g): callable {
    return fn($x) => $f($g($x));
};
echo $compose(fn($n) => $n + 1, fn($n) => $n * 2)(3);
"#,
        ["7"]
    };
}
