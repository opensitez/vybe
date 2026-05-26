use super::helpers::run_prints;

// ── Currying and partial application ─────────────────────────

#[test] fn manual_currying() {
    assert_eq!(run_prints(r#"<?php
function curry(callable $fn): Closure {
    $arity = (new ReflectionFunction($fn))->getNumberOfParameters();
    $args = [];
    $collect = null;
    $collect = function() use ($fn, $arity, &$args, &$collect) {
        $args = array_merge($args, func_get_args());
        return count($args) >= $arity ? $fn(...$args) : $collect;
    };
    return $collect;
}
$add = curry(fn($a,$b) => $a + $b);
$add5 = $add(5);
echo $add5(3) . ',' . $add5(10);
"#), vec!["8,15"]);
}
#[test] fn partial_application_closure() {
    assert_eq!(run_prints(r#"<?php
function partial(callable $fn, mixed ...$partial): Closure {
    return function() use ($fn, $partial) {
        return $fn(...$partial, ...func_get_args());
    };
}
$multiply = fn($a,$b) => $a * $b;
$double = partial($multiply, 2);
$triple = partial($multiply, 3);
echo $double(7) . ',' . $triple(7);
"#), vec!["14,21"]);
}

// ── Function composition ──────────────────────────────────────

#[test] fn compose_two_functions() {
    assert_eq!(run_prints(r#"<?php
function compose(callable $f, callable $g): Closure {
    return fn($x) => $f($g($x));
}
$double = fn($n) => $n * 2;
$inc    = fn($n) => $n + 1;
$doubleInc = compose($double, $inc);
echo $doubleInc(5);
"#), vec!["12"]);
}
#[test] fn pipeline_right_to_left() {
    assert_eq!(run_prints(r#"<?php
function pipe(array $fns): Closure {
    return function($v) use ($fns) {
        return array_reduce($fns, fn($carry, $fn) => $fn($carry), $v);
    };
}
$process = pipe([
    fn($s) => strtolower($s),
    fn($s) => trim($s),
    fn($s) => str_replace(' ', '_', $s),
]);
echo $process('  Hello World  ');
"#), vec!["hello_world"]);
}

// ── Memoization ───────────────────────────────────────────────

#[test] fn memoize_expensive_call() {
    assert_eq!(run_prints(r#"<?php
function memoize(callable $fn): Closure {
    $cache = [];
    return function() use ($fn, &$cache) {
        $key = serialize(func_get_args());
        if (!array_key_exists($key, $cache)) {
            $cache[$key] = $fn(...func_get_args());
        }
        return $cache[$key];
    };
}
$calls = 0;
$expensiveFib = memoize(function(int $n) use (&$calls): int {
    $calls++;
    if ($n <= 1) return $n;
    $fib = 0; $a = 0; $b = 1;
    for ($i = 2; $i <= $n; $i++) { $fib = $a + $b; $a = $b; $b = $fib; }
    return $fib;
});
echo $expensiveFib(10) . ',' . $expensiveFib(10) . ',calls:' . $calls;
"#), vec!["55,55,calls:1"]);
}

// ── array_map / array_filter / array_reduce composition ───────

#[test] fn map_filter_reduce_chain() {
    assert_eq!(run_prints(r#"<?php
$result = array_reduce(
    array_filter(
        array_map(fn($n) => $n * $n, range(1, 10)),
        fn($n) => $n % 2 === 0
    ),
    fn($sum, $n) => $sum + $n,
    0
);
echo $result;
"#), vec!["220"]);
}
#[test] fn transducer_style_processing() {
    assert_eq!(run_prints(r#"<?php
$data = range(1, 20);
$result = array_sum(
    array_slice(
        array_filter($data, fn($n) => $n % 3 === 0),
        0, 4
    )
);
echo $result;
"#), vec!["54"]);
}

// ── Higher order functions ────────────────────────────────────

#[test] fn higher_order_sorter() {
    assert_eq!(run_prints(r#"<?php
function by(callable $key): Closure {
    return fn($a,$b) => $key($a) <=> $key($b);
}
$people = [['n'=>'Charlie','a'=>30],['n'=>'Alice','a'=>25],['n'=>'Bob','a'=>28]];
usort($people, by(fn($p) => $p['a']));
echo implode(',', array_column($people, 'n'));
"#), vec!["Alice,Bob,Charlie"]);
}
#[test] fn flip_arguments() {
    assert_eq!(run_prints(r#"<?php
function flip(callable $fn): Closure {
    return fn($a,$b) => $fn($b,$a);
}
$sub = fn($a,$b) => $a - $b;
echo flip($sub)(3, 10);
"#), vec!["7"]);
}

// ── Closures as state ─────────────────────────────────────────

#[test] fn closure_counter_factory() {
    assert_eq!(run_prints(r#"<?php
function makeCounter(int $start = 0): Closure {
    $n = $start;
    return fn() use (&$n) => ++$n;
}
$c1 = makeCounter();
$c2 = makeCounter(10);
echo $c1() . ',' . $c1() . ',' . $c2() . ',' . $c2();
"#), vec!["1,2,11,12"]);
}
#[test] fn closure_accumulator() {
    assert_eq!(run_prints(r#"<?php
function makeAccumulator(): Closure {
    $total = 0;
    return function(float $n) use (&$total): float { return $total += $n; };
}
$acc = makeAccumulator();
echo $acc(5.0) . ',' . $acc(3.0) . ',' . $acc(2.0);
"#), vec!["5,8,10"]);
}

// ── Recursion patterns ────────────────────────────────────────

#[test] fn mutual_recursion_even_odd() {
    assert_eq!(run_prints(r#"<?php
function isEven(int $n): bool { return $n === 0 ? true : isOdd($n - 1); }
function isOdd(int $n): bool { return $n === 0 ? false : isEven($n - 1); }
echo isEven(4) ? 'even' : 'odd';
echo isOdd(7) ? 'odd' : 'even';
"#), vec!["evenodd"]);
}
#[test] fn trampoline_tail_recursion() {
    assert_eq!(run_prints(r#"<?php
function trampoline(callable $fn): Closure {
    return function() use ($fn) {
        $result = $fn(...func_get_args());
        while (is_callable($result)) $result = $result();
        return $result;
    };
}
$factorial = trampoline(function(int $n, int $acc = 1) use (&$factorial): mixed {
    if ($n <= 1) return $acc;
    return fn() use ($n, $acc) => ($factorial)($n - 1, $n * $acc);
});
echo $factorial(5);
"#), vec!["120"]);
}
