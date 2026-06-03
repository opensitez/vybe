use super::helpers::{compile_ok, run_prints};

// ── Static closures cannot bind $this ─────────────────────────

#[test]
fn static_closure_declared_with_keyword() {
    compile_ok(
        r#"<?php
$fn = static function() {
    return 42;
};
echo $fn();
"#,
    );
}

#[test]
fn static_closure_returns_value() {
    assert_eq!(
        run_prints(
            r#"<?php
$fn = static function(int $x): int {
    return $x * 2;
};
echo $fn(5);
"#
        ),
        vec!["10"]
    );
}

#[test]
fn static_closure_captures_by_value_via_use() {
    assert_eq!(
        run_prints(
            r#"<?php
$factor = 3;
$fn = static function(int $n) use ($factor): int {
    return $n * $factor;
};
echo $fn(4);
"#
        ),
        vec!["12"]
    );
}

#[test]
fn static_closure_cannot_mutate_outer_via_value_capture() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = 10;
$fn = static function() use ($x) {
    $x = 99;
};
$fn();
echo $x;
"#
        ),
        vec!["10"]
    );
}

#[test]
fn static_closure_passed_as_callback_to_array_map() {
    assert_eq!(
        run_prints(
            r#"<?php
$double = static function(int $n): int { return $n * 2; };
$result = array_map($double, [1, 2, 3]);
echo implode(',', $result);
"#
        ),
        vec!["2,4,6"]
    );
}

#[test]
fn static_closure_as_usort_comparator() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = [3, 1, 4, 1, 5];
usort($arr, static function(int $a, int $b): int { return $a <=> $b; });
echo implode(',', $arr);
"#
        ),
        vec!["1,1,3,4,5"]
    );
}

// ── Static arrow functions ────────────────────────────────────

#[test]
fn static_arrow_function_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$double = static fn(int $n): int => $n * 2;
echo $double(7);
"#
        ),
        vec!["14"]
    );
}

#[test]
fn static_arrow_captures_outer_automatically() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = 100;
$add = static fn(int $n): int => $n + $base;
echo $add(5);
"#
        ),
        vec!["105"]
    );
}

#[test]
fn static_arrow_used_in_array_map() {
    assert_eq!(
        run_prints(
            r#"<?php
$offset = 10;
$result = array_map(static fn($x) => $x + $offset, [1, 2, 3]);
echo implode(',', $result);
"#
        ),
        vec!["11,12,13"]
    );
}

// ── Closure::bind with static ─────────────────────────────────

#[test]
fn closure_bind_to_new_object() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    private int $count = 0;
}
$increment = Closure::bind(function() {
    $this->count++;
    return $this->count;
}, new Counter(), Counter::class);
echo $increment();
"#
        ),
        vec!["1"]
    );
}

#[test]
fn closure_bind_with_null_scope_static() {
    compile_ok(
        r#"<?php
$fn = static function() { return "static"; };
$bound = Closure::bind($fn, null, null);
echo $bound();
"#,
    );
}

#[test]
fn closure_bindto_instance_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box {
    private string $label = "box";
}
$getLabel = function() { return $this->label; };
$b = new Box();
$bound = Closure::bind($getLabel, $b, Box::class);
echo $bound();
"#
        ),
        vec!["box"]
    );
}

// ── Closure::fromCallable ─────────────────────────────────────

#[test]
fn closure_from_callable_named_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function double(int $n): int { return $n * 2; }
$fn = Closure::fromCallable('double');
echo $fn(6);
"#
        ),
        vec!["12"]
    );
}

#[test]
fn closure_from_callable_builtin() {
    assert_eq!(
        run_prints(
            r#"<?php
$len = Closure::fromCallable('strlen');
echo $len("hello");
"#
        ),
        vec!["5"]
    );
}

#[test]
fn closure_from_callable_instance_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class Calc {
    public function square(int $n): int { return $n ** 2; }
}
$c = new Calc();
$fn = Closure::fromCallable([$c, 'square']);
echo $fn(4);
"#
        ),
        vec!["16"]
    );
}

// ── Recursive closures via reference capture ──────────────────

#[test]
fn recursive_closure_via_reference() {
    assert_eq!(
        run_prints(
            r#"<?php
$factorial = null;
$factorial = function(int $n) use (&$factorial): int {
    return $n <= 1 ? 1 : $n * $factorial($n - 1);
};
echo $factorial(5);
"#
        ),
        vec!["120"]
    );
}

#[test]
fn recursive_closure_fibonacci() {
    assert_eq!(
        run_prints(
            r#"<?php
$fib = null;
$fib = function(int $n) use (&$fib): int {
    if ($n <= 1) return $n;
    return $fib($n - 1) + $fib($n - 2);
};
echo $fib(7);
"#
        ),
        vec!["13"]
    );
}

// ── Closure memoization pattern ───────────────────────────────

#[test]
fn memoize_closure_with_static_cache() {
    assert_eq!(
        run_prints(
            r#"<?php
function memoize(callable $fn): Closure {
    $cache = [];
    return function() use ($fn, &$cache) {
        $args = func_get_args();
        $key = serialize($args);
        if (!isset($cache[$key])) {
            $cache[$key] = $fn(...$args);
        }
        return $cache[$key];
    };
}
$expensive = memoize(function(int $n): int { return $n * $n; });
echo $expensive(4);
echo $expensive(4);
"#
        ),
        vec!["16", "16"]
    );
}

// ── Partial application via closure ──────────────────────────

#[test]
fn partial_application_via_closure() {
    assert_eq!(
        run_prints(
            r#"<?php
function partial(callable $fn, mixed ...$partial): Closure {
    return function() use ($fn, $partial) {
        $args = array_merge($partial, func_get_args());
        return $fn(...$args);
    };
}
$add = fn(int $a, int $b): int => $a + $b;
$add5 = partial($add, 5);
echo $add5(3);
"#
        ),
        vec!["8"]
    );
}

// ── Closure type hints ────────────────────────────────────────

#[test]
fn closure_type_hint_in_parameter() {
    assert_eq!(
        run_prints(
            r#"<?php
function apply(Closure $fn, int $val): int {
    return $fn($val);
}
echo apply(fn($x) => $x + 10, 5);
"#
        ),
        vec!["15"]
    );
}

#[test]
fn callable_type_hint_accepts_closure() {
    assert_eq!(
        run_prints(
            r#"<?php
function transform(callable $fn, array $items): array {
    return array_map($fn, $items);
}
$result = transform(static fn($x) => $x ** 2, [1, 2, 3, 4]);
echo implode(',', $result);
"#
        ),
        vec!["1,4,9,16"]
    );
}

// ── Static closure in class context ──────────────────────────

#[test]
fn static_closure_returned_from_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class Factory {
    public static function multiplier(int $factor): Closure {
        return static function(int $n) use ($factor): int {
            return $n * $factor;
        };
    }
}
$triple = Factory::multiplier(3);
echo $triple(7);
"#
        ),
        vec!["21"]
    );
}

#[test]
fn closure_used_as_default_property_initializer_workaround() {
    assert_eq!(
        run_prints(
            r#"<?php
class Pipeline {
    private array $stages = [];
    public function pipe(Closure $stage): static {
        $this->stages[] = $stage;
        return $this;
    }
    public function run(mixed $payload): mixed {
        foreach ($this->stages as $stage) {
            $payload = $stage($payload);
        }
        return $payload;
    }
}
$result = (new Pipeline())
    ->pipe(static fn($x) => $x * 2)
    ->pipe(static fn($x) => $x + 1)
    ->run(5);
echo $result;
"#
        ),
        vec!["11"]
    );
}
