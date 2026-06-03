use super::helpers::{compile_ok, run_prints};

// ── First-class callable syntax (PHP 8.1+) ───────────────────────

#[test]
fn strlen_stored_in_variable() {
    assert_eq!(
        run_prints(
            r#"<?php
$fn = strlen(...);
echo $fn('hello') . "\n";
echo $fn('') . "\n";
"#
        ),
        vec!["5", "0"]
    );
}

#[test]
fn strtoupper_passed_to_array_map() {
    assert_eq!(
        run_prints(
            r#"<?php
$fn = strtoupper(...);
$result = array_map($fn, ['hello', 'world', 'php']);
echo implode(',', $result) . "\n";
"#
        ),
        vec!["HELLO,WORLD,PHP"]
    );
}

#[test]
fn trim_in_array_filter_usage() {
    assert_eq!(
        run_prints(
            r#"<?php
$fn = trim(...);
$items = array_map($fn, ['  hello  ', ' world ', 'php']);
echo implode('|', $items) . "\n";
"#
        ),
        vec!["hello|world|php"]
    );
}

#[test]
fn intval_in_array_map() {
    assert_eq!(
        run_prints(
            r#"<?php
$fn = intval(...);
$result = array_map($fn, ['1', '2', '3', '4']);
echo array_sum($result) . "\n";
"#
        ),
        vec!["10"]
    );
}

#[test]
fn user_function_as_first_class_callable() {
    assert_eq!(
        run_prints(
            r#"<?php
function double(int $n): int { return $n * 2; }
$fn = double(...);
echo $fn(5) . "\n";
$results = array_map($fn, [1, 2, 3]);
echo implode(',', $results) . "\n";
"#
        ),
        vec!["10", "2,4,6"]
    );
}

#[test]
fn static_method_first_class_callable() {
    assert_eq!(
        run_prints(
            r#"<?php
class MathUtils {
    public static function square(int $n): int { return $n * $n; }
    public static function cube(int $n): int { return $n * $n * $n; }
}
$sq = MathUtils::square(...);
$cu = MathUtils::cube(...);
echo $sq(4) . "\n";
echo $cu(3) . "\n";
$result = array_map($sq, [1, 2, 3, 4]);
echo implode(',', $result) . "\n";
"#
        ),
        vec!["16", "27", "1,4,9,16"]
    );
}

#[test]
fn instance_method_first_class_callable() {
    assert_eq!(
        run_prints(
            r#"<?php
class Multiplier {
    public function __construct(private int $factor) {}
    public function multiply(int $n): int { return $n * $this->factor; }
}
$m = new Multiplier(3);
$fn = $m->multiply(...);
echo $fn(4) . "\n";
$result = array_map($fn, [1, 2, 3]);
echo implode(',', $result) . "\n";
"#
        ),
        vec!["12", "3,6,9"]
    );
}

#[test]
fn closure_already_first_class() {
    assert_eq!(
        run_prints(
            r#"<?php
$add = function(int $a, int $b): int { return $a + $b; };
$result = array_map(fn($x) => $add($x, 10), [1, 2, 3]);
echo implode(',', $result) . "\n";
"#
        ),
        vec!["11,12,13"]
    );
}

#[test]
fn first_class_callable_in_usort() {
    assert_eq!(
        run_prints(
            r#"<?php
function compareDesc(int $a, int $b): int { return $b <=> $a; }
$arr = [3, 1, 4, 1, 5, 9, 2, 6];
usort($arr, compareDesc(...));
echo implode(',', $arr) . "\n";
"#
        ),
        vec!["9,6,5,4,3,2,1,1"]
    );
}

#[test]
fn first_class_callable_passed_to_accepting_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function applyAll(array $fns, $value) {
    return array_reduce($fns, fn($carry, $fn) => $fn($carry), $value);
}
$result = applyAll([strtoupper(...), trim(...), strrev(...)], '  hello  ');
echo $result . "\n";
"#
        ),
        vec!["OLLEH"]
    );
}

#[test]
fn is_callable_on_first_class_callable() {
    assert_eq!(
        run_prints(
            r#"<?php
$fn = strlen(...);
echo is_callable($fn) ? 'callable' : 'not callable';
echo "\n";
$staticFn = DateTime::createFromFormat(...);
echo is_callable($staticFn) ? 'callable' : 'not callable';
echo "\n";
"#
        ),
        vec!["callable", "callable"]
    );
}

#[test]
fn first_class_callable_with_call_user_func() {
    assert_eq!(
        run_prints(
            r#"<?php
$fn = strtolower(...);
echo call_user_func($fn, 'HELLO WORLD') . "\n";
"#
        ),
        vec!["hello world"]
    );
}

#[test]
fn first_class_callable_composition_wrap() {
    assert_eq!(
        run_prints(
            r#"<?php
function compose(callable ...$fns): callable {
    return function($x) use ($fns) {
        return array_reduce(
            array_reverse($fns),
            fn($carry, $fn) => $fn($carry),
            $x
        );
    };
}
$transform = compose(strtoupper(...), trim(...));
echo $transform('  hello  ') . "\n";
"#
        ),
        vec!["HELLO"]
    );
}

#[test]
fn partial_application_via_closure_wrapping() {
    assert_eq!(
        run_prints(
            r#"<?php
function partial(callable $fn, ...$partial): callable {
    return function() use ($fn, $partial) {
        $args = array_merge($partial, func_get_args());
        return $fn(...$args);
    };
}
function add(int $a, int $b): int { return $a + $b; }
$add5 = partial(add(...), 5);
echo $add5(3) . "\n";
echo $add5(10) . "\n";
$result = array_map($add5, [1, 2, 3]);
echo implode(',', $result) . "\n";
"#
        ),
        vec!["8", "15", "6,7,8"]
    );
}

#[test]
fn static_closure_as_first_class() {
    assert_eq!(
        run_prints(
            r#"<?php
$fn = static function(int $n): int { return $n * $n; };
echo $fn(5) . "\n";
$result = array_map($fn, [2, 3, 4]);
echo implode(',', $result) . "\n";
"#
        ),
        vec!["25", "4,9,16"]
    );
}

#[test]
fn chaining_callables_in_pipeline() {
    assert_eq!(
        run_prints(
            r#"<?php
function pipeline($value, callable ...$fns) {
    foreach ($fns as $fn) $value = $fn($value);
    return $value;
}
$result = pipeline(
    '  PHP is Great  ',
    trim(...),
    strtolower(...),
    fn($s) => str_replace(' ', '_', $s)
);
echo $result . "\n";
"#
        ),
        vec!["php_is_great"]
    );
}

#[test]
fn builtin_abs_as_first_class() {
    assert_eq!(
        run_prints(
            r#"<?php
$fn = abs(...);
$result = array_map($fn, [-3, -1, 0, 2, -5]);
echo implode(',', $result) . "\n";
"#
        ),
        vec!["3,1,0,2,5"]
    );
}

#[test]
fn builtin_round_in_array_map() {
    assert_eq!(
        run_prints(
            r#"<?php
$fn = round(...);
$result = array_map($fn, [1.4, 1.5, 2.6, 3.1, -1.5]);
echo implode(',', $result) . "\n";
"#
        ),
        vec!["1,2,3,3,-2"]
    );
}

#[test]
fn first_class_callable_preserving_this() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    private int $count = 0;
    public function increment(int $by = 1): void { $this->count += $by; }
    public function getCount(): int { return $this->count; }
}
$counter = new Counter();
$inc = $counter->increment(...);
$get = $counter->getCount(...);
$inc(5);
$inc(3);
echo $get() . "\n";
array_map($inc, [1, 1, 1]);
echo $get() . "\n";
"#
        ),
        vec!["8", "11"]
    );
}

#[test]
fn first_class_callable_stored_in_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$ops = [
    'upper' => strtoupper(...),
    'lower' => strtolower(...),
    'rev'   => strrev(...),
];
echo $ops['upper']('hello') . "\n";
echo $ops['lower']('WORLD') . "\n";
echo $ops['rev']('abcde') . "\n";
"#
        ),
        vec!["HELLO", "world", "edcba"]
    );
}

#[test]
fn first_class_callable_in_match() {
    assert_eq!(
        run_prints(
            r#"<?php
function getTransform(string $name): callable {
    return match($name) {
        'upper' => strtoupper(...),
        'lower' => strtolower(...),
        'trim'  => trim(...),
        default => fn($s) => $s,
    };
}
echo getTransform('upper')('hello') . "\n";
echo getTransform('lower')('WORLD') . "\n";
echo getTransform('trim')('  php  ') . "\n";
"#
        ),
        vec!["HELLO", "world", "php"]
    );
}

#[test]
fn callable_type_hint_accepts_first_class() {
    assert_eq!(
        run_prints(
            r#"<?php
function applyToAll(callable $fn, array $items): array {
    return array_map($fn, $items);
}
$result = applyToAll(strtoupper(...), ['foo', 'bar', 'baz']);
echo implode(',', $result) . "\n";
"#
        ),
        vec!["FOO,BAR,BAZ"]
    );
}

#[test]
fn first_class_callable_of_variadic_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function joinWith(string $sep, string ...$parts): string {
    return implode($sep, $parts);
}
$fn = joinWith(...);
echo $fn('-', 'a', 'b', 'c') . "\n";
echo $fn('|', 'x', 'y') . "\n";
"#
        ),
        vec!["a-b-c", "x|y"]
    );
}

#[test]
fn memoize_via_first_class_callable() {
    assert_eq!(
        run_prints(
            r#"<?php
function memoize(callable $fn): callable {
    $cache = [];
    return function() use ($fn, &$cache) {
        $key = serialize(func_get_args());
        if (!array_key_exists($key, $cache)) {
            $cache[$key] = $fn(...func_get_args());
        }
        return $cache[$key];
    };
}
$callCount = 0;
function expensiveCompute(int $n) use (&$callCount): int {
    $callCount++;
    return $n * $n;
}
$memoized = memoize(expensiveCompute(...));
echo $memoized(4) . "\n";
echo $memoized(4) . "\n";
echo $memoized(5) . "\n";
echo $callCount . "\n";
"#
        ),
        vec!["16", "16", "25", "2"]
    );
}

#[test]
fn first_class_callable_from_array_map() {
    assert_eq!(
        run_prints(
            r#"<?php
$numbers = range(1, 5);
$squared = array_map(fn($n) => $n ** 2, $numbers);
$toStr = strval(...);
$strings = array_map($toStr, $squared);
echo implode(',', $strings) . "\n";
"#
        ),
        vec!["1,4,9,16,25"]
    );
}

#[test]
fn first_class_callable_from_parent_method() {
    compile_ok(
        r#"<?php
class Base {
    public function transform(string $s): string { return strtoupper($s); }
}
class Child extends Base {
    public function getParentTransform(): callable {
        return parent::transform(...);
    }
}
$child = new Child();
$fn = $child->getParentTransform();
echo $fn('hello');
"#,
    );
}
