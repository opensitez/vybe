use super::helpers::compile_ok;

// ── array_map with arrow function ────────────────────────────────
#[test]
fn array_map_arrow_function() {
    compile_ok(r#"<?php
$nums = [1, 2, 3, 4, 5];
$doubled = array_map(fn($n) => $n * 2, $nums);
echo implode(',', $doubled);
"#);
}

// ── array_filter with arrow function ─────────────────────────────
#[test]
fn array_filter_arrow_function() {
    compile_ok(r#"<?php
$nums = [1, 2, 3, 4, 5, 6, 7, 8];
$evens = array_filter($nums, fn($n) => $n % 2 === 0);
echo implode(',', array_values($evens));
"#);
}

// ── array_reduce building a sum ───────────────────────────────────
#[test]
fn array_reduce_sum() {
    compile_ok(r#"<?php
$nums = [1, 2, 3, 4, 5];
$sum = array_reduce($nums, fn($carry, $n) => $carry + $n, 0);
echo $sum;
"#);
}

// ── array_reduce building nested structure ────────────────────────
#[test]
fn array_reduce_build_map() {
    compile_ok(r#"<?php
$pairs = [['k' => 'a', 'v' => 1], ['k' => 'b', 'v' => 2], ['k' => 'c', 'v' => 3]];
$map = array_reduce($pairs, function($carry, $item) {
    $carry[$item['k']] = $item['v'];
    return $carry;
}, []);
echo $map['b'];
echo count($map);
"#);
}

// ── usort with spaceship operator ────────────────────────────────
#[test]
fn usort_spaceship_operator() {
    compile_ok(r#"<?php
$words = ['banana', 'apple', 'cherry', 'date'];
usort($words, fn($a, $b) => $a <=> $b);
echo implode(',', $words);
"#);
}

// ── Function composition (two callables) ─────────────────────────
#[test]
fn function_composition() {
    compile_ok(r#"<?php
function compose(callable $f, callable $g): callable {
    return fn($x) => $f($g($x));
}
$trim   = fn($s) => trim($s);
$upper  = fn($s) => strtoupper($s);
$clean  = compose($upper, $trim);
echo $clean('  hello world  ');
"#);
}

// ── Partial application capturing some args ───────────────────────
#[test]
fn partial_application() {
    compile_ok(r#"<?php
function multiply(int $a, int $b): int { return $a * $b; }
function partial(callable $fn, ...$partialArgs): callable {
    return fn(...$rest) => $fn(...$partialArgs, ...$rest);
}
$double = partial('multiply', 2);
$triple = partial('multiply', 3);
echo $double(5);
echo $triple(5);
"#);
}

// ── Memoization with static cache array ──────────────────────────
#[test]
fn memoization_static_cache() {
    compile_ok(r#"<?php
function memoize(callable $fn): callable {
    $cache = [];
    return function() use ($fn, &$cache) {
        $key = serialize(func_get_args());
        if (!isset($cache[$key])) {
            $cache[$key] = $fn(...func_get_args());
        }
        return $cache[$key];
    };
}
$expensiveAdd = memoize(fn($a, $b) => $a + $b);
echo $expensiveAdd(3, 4);
echo $expensiveAdd(3, 4);
"#);
}

// ── Currying — function returning function ────────────────────────
#[test]
fn currying_function_returning_function() {
    compile_ok(r#"<?php
function curry(callable $fn): callable {
    $arity = (new ReflectionFunction(Closure::fromCallable($fn)))->getNumberOfParameters();
    $accumulate = function(array $args) use ($fn, $arity, &$accumulate): mixed {
        if (count($args) >= $arity) {
            return $fn(...$args);
        }
        return fn(...$more) => $accumulate(array_merge($args, $more));
    };
    return fn(...$args) => $accumulate($args);
}
$add = curry(fn($a, $b) => $a + $b);
$add5 = $add(5);
echo $add5(3);
echo $add(10)(20);
"#);
}

// ── Pipeline — value through chain of transforms ──────────────────
#[test]
fn pipeline_value_transforms() {
    compile_ok(r#"<?php
function pipeline($value, callable ...$fns) {
    return array_reduce($fns, fn($carry, $fn) => $fn($carry), $value);
}
$result = pipeline(
    '  hello world  ',
    'trim',
    'strtoupper',
    fn($s) => str_replace(' ', '-', $s)
);
echo $result;
"#);
}

// ── array_map returning objects ───────────────────────────────────
#[test]
fn array_map_returning_objects() {
    compile_ok(r#"<?php
class Point {
    public function __construct(public int $x, public int $y) {}
}
$coords = [[1, 2], [3, 4], [5, 6]];
$points = array_map(fn($c) => new Point($c[0], $c[1]), $coords);
echo $points[1]->x . ',' . $points[1]->y;
echo count($points);
"#);
}

// ── Recursive array flatten using array_reduce ────────────────────
#[test]
fn recursive_flatten_with_reduce() {
    compile_ok(r#"<?php
function flatten(array $arr): array {
    return array_reduce($arr, function($carry, $item) {
        if (is_array($item)) {
            return array_merge($carry, flatten($item));
        }
        $carry[] = $item;
        return $carry;
    }, []);
}
$nested = [1, [2, 3], [4, [5, 6]]];
$flat = flatten($nested);
echo implode(',', $flat);
"#);
}

// ── group_by implemented with array_reduce ────────────────────────
#[test]
fn group_by_with_reduce() {
    compile_ok(r#"<?php
function groupBy(array $items, callable $keyFn): array {
    return array_reduce($items, function($groups, $item) use ($keyFn) {
        $key = $keyFn($item);
        $groups[$key][] = $item;
        return $groups;
    }, []);
}
$people = [
    ['name' => 'Alice', 'dept' => 'eng'],
    ['name' => 'Bob',   'dept' => 'hr'],
    ['name' => 'Carol', 'dept' => 'eng'],
];
$grouped = groupBy($people, fn($p) => $p['dept']);
echo count($grouped['eng']);
echo $grouped['hr'][0]['name'];
"#);
}

// ── partition implemented with array_reduce ───────────────────────
#[test]
fn partition_with_reduce() {
    compile_ok(r#"<?php
function partition(array $items, callable $pred): array {
    return array_reduce($items, function($parts, $item) use ($pred) {
        $parts[$pred($item) ? 0 : 1][] = $item;
        return $parts;
    }, [[], []]);
}
$nums = [1, 2, 3, 4, 5, 6, 7, 8];
[$evens, $odds] = partition($nums, fn($n) => $n % 2 === 0);
echo implode(',', $evens);
echo implode(',', $odds);
"#);
}

// ── zip two arrays using array_map(null, ...) ─────────────────────
#[test]
fn zip_arrays_with_null_map() {
    compile_ok(r#"<?php
$keys   = ['a', 'b', 'c'];
$values = [1,   2,   3  ];
$zipped = array_map(null, $keys, $values);
foreach ($zipped as [$k, $v]) {
    echo "$k=$v ";
}
"#);
}

// ── take_while using foreach + break ─────────────────────────────
#[test]
fn take_while_foreach_break() {
    compile_ok(r#"<?php
function takeWhile(array $items, callable $pred): array {
    $result = [];
    foreach ($items as $item) {
        if (!$pred($item)) break;
        $result[] = $item;
    }
    return $result;
}
$nums = [1, 2, 3, 4, 5, 1, 2];
$taken = takeWhile($nums, fn($n) => $n < 4);
echo implode(',', $taken);
"#);
}

// ── drop_while using foreach + flag ──────────────────────────────
#[test]
fn drop_while_foreach_flag() {
    compile_ok(r#"<?php
function dropWhile(array $items, callable $pred): array {
    $dropping = true;
    $result   = [];
    foreach ($items as $item) {
        if ($dropping && $pred($item)) continue;
        $dropping = false;
        $result[] = $item;
    }
    return $result;
}
$nums = [1, 2, 3, 4, 5];
$dropped = dropWhile($nums, fn($n) => $n < 3);
echo implode(',', $dropped);
"#);
}

// ── Lazy evaluation with generators (yield) ───────────────────────
#[test]
fn generator_lazy_range() {
    compile_ok(r#"<?php
function lazyRange(int $start, int $end): Generator {
    for ($i = $start; $i <= $end; $i++) {
        yield $i;
    }
}
$gen = lazyRange(1, 5);
foreach ($gen as $n) {
    echo $n;
}
"#);
}

// ── Function returning closure over state ─────────────────────────
#[test]
fn function_returns_closure_over_state() {
    compile_ok(r#"<?php
function makeAccumulator(int $initial = 0): callable {
    $total = $initial;
    return function(int $n) use (&$total): int {
        $total += $n;
        return $total;
    };
}
$acc = makeAccumulator(10);
echo $acc(5);
echo $acc(3);
echo $acc(2);
"#);
}

// ── Higher order function accepting callable typehint ─────────────
#[test]
fn higher_order_callable_typehint() {
    compile_ok(r#"<?php
function applyTwice(callable $fn, mixed $value): mixed {
    return $fn($fn($value));
}
$addTen   = fn($x) => $x + 10;
$toUpper  = fn($s) => strtoupper($s);
echo applyTwice($addTen, 5);
echo applyTwice($toUpper, 'hi');
"#);
}
