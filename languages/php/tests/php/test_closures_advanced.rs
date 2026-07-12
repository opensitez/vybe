use super::helpers::compile_ok;

// ── Closure::bind ─────────────────────────────────────────────

#[test]
fn closure_bind_basic() {
    compile_ok(
        r#"<?php
class Foo { private int $x = 42; }
$fn = Closure::bind(function() { return $this->x; }, new Foo(), 'Foo');
echo $fn();
"#,
    );
}

#[test]
fn closure_bind_change_object() {
    compile_ok(
        r#"<?php
class Counter { private int $count = 0; }
$inc = Closure::bind(function() { $this->count++; return $this->count; }, new Counter(), Counter::class);
echo $inc();
echo $inc();
"#,
    );
}

#[test]
fn closure_bind_static_context() {
    compile_ok(
        r#"<?php
class Registry { private static array $items = []; }
$add = Closure::bind(
    static function(string $k, mixed $v) { static::$items[$k] = $v; },
    null,
    Registry::class
);
$add('key', 'value');
$get = Closure::bind(
    static function(string $k) { return static::$items[$k] ?? null; },
    null,
    Registry::class
);
echo $get('key');
"#,
    );
}

#[test]
fn closure_bind_null_object_static() {
    compile_ok(
        r#"<?php
class Config {
    private static string $env = 'production';
    public static function getEnv(): string { return static::$env; }
}
$reader = Closure::bind(
    static function() { return static::$env; },
    null,
    Config::class
);
echo $reader();
"#,
    );
}

// ── Closure::bindTo ───────────────────────────────────────────

#[test]
fn closure_bind_to_basic() {
    compile_ok(
        r#"<?php
class Box { private int $size = 10; }
$fn = function() { return $this->size; };
$bound = $fn->bindTo(new Box(), Box::class);
echo $bound();
"#,
    );
}

#[test]
fn closure_bind_to_different_instance() {
    compile_ok(
        r#"<?php
class Node { private string $label; public function __construct(string $l) { $this->label = $l; } }
$getLabel = function() { return $this->label; };
$a = new Node('alpha');
$b = new Node('beta');
$fa = $getLabel->bindTo($a, Node::class);
$fb = $getLabel->bindTo($b, Node::class);
echo $fa() . ',' . $fb();
"#,
    );
}

#[test]
fn closure_bind_to_fluent() {
    compile_ok(
        r#"<?php
class Builder {
    private array $parts = [];
    public function add(string $part): static { $this->parts[] = $part; return $this; }
    public function build(): string { return implode('-', $this->parts); }
}
$addPart = (function(string $p) { $this->parts[] = $p; return $this; })->bindTo(new Builder(), Builder::class);
$b = new Builder();
$b->add('a')->add('b');
echo $b->build();
"#,
    );
}

// ── Closure::fromCallable ─────────────────────────────────────

#[test]
fn closure_from_callable_function() {
    compile_ok(
        r#"<?php
function double(int $n): int { return $n * 2; }
$fn = Closure::fromCallable('double');
echo $fn(21);
"#,
    );
}

#[test]
fn closure_from_callable_method() {
    compile_ok(
        r#"<?php
class Math {
    public function square(int $n): int { return $n * $n; }
}
$m = new Math();
$fn = Closure::fromCallable([$m, 'square']);
echo $fn(9);
"#,
    );
}

#[test]
fn closure_from_callable_static_method() {
    compile_ok(
        r#"<?php
class Util {
    public static function triple(int $n): int { return $n * 3; }
}
$fn = Closure::fromCallable(['Util', 'triple']);
echo $fn(7);
"#,
    );
}

#[test]
fn closure_from_callable_builtin() {
    compile_ok(
        r#"<?php
$upper = Closure::fromCallable('strtoupper');
$len   = Closure::fromCallable('strlen');
echo $upper('hello');
echo $len('hello');
"#,
    );
}

#[test]
fn closure_from_callable_compose() {
    compile_ok(
        r#"<?php
function compose(callable ...$fns): Closure {
    return function($v) use ($fns) {
        return array_reduce(
            array_reverse($fns),
            fn($carry, $fn) => $fn($carry),
            $v
        );
    };
}
$process = compose(
    Closure::fromCallable('strtoupper'),
    Closure::fromCallable('trim')
);
echo $process("  hello world  ");
"#,
    );
}

// ── Static closures ───────────────────────────────────────────

#[test]
fn static_closure_basic() {
    compile_ok(
        r#"<?php
$fn = static function(int $n): int { return $n * 2; };
echo $fn(21);
"#,
    );
}

#[test]
fn static_closure_no_this() {
    compile_ok(
        r#"<?php
class Foo {
    public int $x = 5;
    public function getStatic(): Closure {
        return static function() {
            // $this is not available here
            return 'static closure';
        };
    }
}
$f = new Foo();
echo $f->getStatic()();
"#,
    );
}

#[test]
fn static_arrow_function() {
    compile_ok(
        r#"<?php
$fn = static fn(int $a, int $b) => $a + $b;
echo $fn(10, 32);
"#,
    );
}

#[test]
fn static_closure_in_array_map() {
    compile_ok(
        r#"<?php
$nums = [1, 2, 3, 4, 5];
$squares = array_map(static fn(int $n) => $n ** 2, $nums);
echo implode(',', $squares);
"#,
    );
}

// ── Closure as first-class type ───────────────────────────────

#[test]
fn closure_type_hint() {
    compile_ok(
        r#"<?php
function apply(Closure $fn, int $v): int { return $fn($v); }
echo apply(fn($x) => $x * 3, 14);
"#,
    );
}

#[test]
fn closure_stored_in_property() {
    compile_ok(
        r#"<?php
class Handler {
    private Closure $callback;
    public function __construct(Closure $cb) { $this->callback = $cb; }
    public function handle(string $input): string { return ($this->callback)($input); }
}
$h = new Handler(fn($s) => strtoupper(trim($s)));
echo $h->handle("  hello  ");
"#,
    );
}

#[test]
fn closure_returned_from_function() {
    compile_ok(
        r#"<?php
function multiplier(int $factor): Closure {
    return fn(int $n) => $n * $factor;
}
$double = multiplier(2);
$triple = multiplier(3);
echo $double(5) . ',' . $triple(5);
"#,
    );
}

#[test]
fn closure_memoize() {
    compile_ok(
        r#"<?php
function memoize(Closure $fn): Closure {
    $cache = [];
    return function() use ($fn, &$cache) {
        $key = serialize(func_get_args());
        if (!array_key_exists($key, $cache)) {
            $cache[$key] = $fn(...func_get_args());
        }
        return $cache[$key];
    };
}
$fib = memoize(function(int $n) use (&$fib): int {
    if ($n <= 1) return $n;
    return $fib($n - 1) + $fib($n - 2);
});
echo $fib(10);
"#,
    );
}

// ── Partial application via closures ──────────────────────────

#[test]
fn partial_application() {
    compile_ok(
        r#"<?php
function partial(callable $fn, mixed ...$partialArgs): Closure {
    return function() use ($fn, $partialArgs) {
        $args = array_merge($partialArgs, func_get_args());
        return $fn(...$args);
    };
}
function add(int $a, int $b, int $c): int { return $a + $b + $c; }
$add10 = partial('add', 10);
$add10and20 = partial('add', 10, 20);
echo $add10(5, 3);
echo $add10and20(7);
"#,
    );
}

#[test]
fn currying_with_closures() {
    compile_ok(
        r#"<?php
function curry(callable $fn): Closure {
    $arity = (new ReflectionFunction(Closure::fromCallable($fn)))->getNumberOfParameters();
    $accumulate = function(array $args) use ($fn, $arity, &$accumulate): mixed {
        if (count($args) >= $arity) return $fn(...$args);
        return function() use ($args, $accumulate) {
            return $accumulate(array_merge($args, func_get_args()));
        };
    };
    return function() use ($accumulate) { return $accumulate(func_get_args()); };
}
$add = curry(fn(int $a, int $b, int $c) => $a + $b + $c);
echo $add(1)(2)(3);
"#,
    );
}

// ── Closure::call (PHP 7.0+) ──────────────────────────────────

#[test]
fn closure_call_method() {
    compile_ok(
        r#"<?php
class Secret { private string $value = 'hidden'; }
$fn = function() { return $this->value; };
echo $fn->call(new Secret());
"#,
    );
}

#[test]
fn closure_call_with_args() {
    compile_ok(
        r#"<?php
class Multiplier { private int $factor = 3; }
$fn = function(int $n): int { return $this->factor * $n; };
echo $fn->call(new Multiplier(), 7);
"#,
    );
}
