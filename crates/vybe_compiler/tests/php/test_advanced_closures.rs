use super::helpers::compile_ok;

// ── Closure as default parameter ─────────────────────────────

#[test] fn closure_as_default_parameter_value() {
    compile_ok(r#"<?php
function transform(array $data, callable $fn = null): array {
    $fn ??= fn($x) => $x;
    return array_map($fn, $data);
}
$result = transform([1, 2, 3]);
echo count($result);
"#);
}

// ── Closure stored in object property ────────────────────────

#[test] fn closure_stored_in_object_property_called() {
    compile_ok(r#"<?php
class Handler {
    public Closure $onEvent;
    public function __construct() {
        $this->onEvent = function(string $name): string { return 'handled: ' . $name; };
    }
}
$h = new Handler();
echo ($h->onEvent)('click');
"#);
}

// ── Closure modifying external array by reference ─────────────

#[test] fn closure_modifies_external_array_by_ref() {
    compile_ok(r#"<?php
$items = [];
$collect = function(mixed $v) use (&$items): void { $items[] = $v; };
$collect('a');
$collect('b');
$collect('c');
echo count($items);
"#);
}

// ── Multiple use-by-reference ─────────────────────────────────

#[test] fn closure_multiple_use_by_ref() {
    compile_ok(r#"<?php
$sum = 0;
$count = 0;
$record = function(int $n) use (&$sum, &$count): void {
    $sum += $n;
    $count++;
};
$record(10);
$record(20);
$record(30);
echo $sum;
echo $count;
"#);
}

// ── Closure returned from method ─────────────────────────────

#[test] fn closure_returned_from_method_bound_to_object() {
    compile_ok(r#"<?php
class Greeter {
    private string $prefix;
    public function __construct(string $prefix) { $this->prefix = $prefix; }
    public function makeGreeter(): Closure {
        return function(string $name): string { return $this->prefix . $name; };
    }
}
$g = new Greeter('Hello, ');
$fn = $g->makeGreeter();
echo $fn('World');
"#);
}

// ── Array of closures ─────────────────────────────────────────

#[test] fn array_of_closures_iterate_and_call() {
    compile_ok(r#"<?php
$fns = [
    fn(int $x): int => $x + 1,
    fn(int $x): int => $x * 2,
    fn(int $x): int => $x - 3,
];
$val = 10;
foreach ($fns as $fn) {
    $val = $fn($val);
}
echo $val;
"#);
}

// ── Closure composition ───────────────────────────────────────

#[test] fn closure_composition_f_of_g() {
    compile_ok(r#"<?php
function compose(callable $f, callable $g): Closure {
    return fn(mixed $x) => $f($g($x));
}
$double = fn(int $x) => $x * 2;
$addTen = fn(int $x) => $x + 10;
$doubleThenAdd = compose($addTen, $double);
echo $doubleThenAdd(5);
"#);
}

// ── Closure factory returning closure ────────────────────────

#[test] fn closure_factory_returns_closure() {
    compile_ok(r#"<?php
$makeAdder = function(int $base): Closure {
    return function(int $x) use ($base): int { return $base + $x; };
};
$add100 = $makeAdder(100);
echo $add100(42);
"#);
}

// ── Recursive closure via use(&$fn) ──────────────────────────

#[test] fn recursive_closure_via_use_ref() {
    compile_ok(r#"<?php
$factorial = null;
$factorial = function(int $n) use (&$factorial): int {
    return $n <= 1 ? 1 : $n * $factorial($n - 1);
};
echo $factorial(5);
"#);
}

// ── Closure as usort comparator ───────────────────────────────

#[test] fn closure_complex_usort_comparator() {
    compile_ok(r#"<?php
$people = [
    ['name' => 'Zara', 'age' => 25],
    ['name' => 'Alice', 'age' => 30],
    ['name' => 'Bob', 'age' => 25],
];
usort($people, function(array $a, array $b): int {
    if ($a['age'] !== $b['age']) return $a['age'] <=> $b['age'];
    return $a['name'] <=> $b['name'];
});
echo $people[0]['name'];
"#);
}

// ── Closure in match arm ──────────────────────────────────────

#[test] fn closure_in_match_expression_arm() {
    compile_ok(r#"<?php
$op = 'double';
$fn = match ($op) {
    'double' => fn(int $x) => $x * 2,
    'square' => fn(int $x) => $x * $x,
    default  => fn(int $x) => $x,
};
echo $fn(7);
"#);
}

// ── Closure in ternary ────────────────────────────────────────

#[test] fn closure_in_ternary_expression() {
    compile_ok(r#"<?php
$flag = true;
$transform = $flag ? fn(int $x) => $x * 10 : fn(int $x) => $x;
echo $transform(5);
"#);
}

// ── Static closure ────────────────────────────────────────────

#[test] fn static_closure_no_this_binding() {
    compile_ok(r#"<?php
class Widget {
    public static function getTransformer(): Closure {
        return static function(int $n): int { return $n ** 2; };
    }
}
$fn = Widget::getTransformer();
echo $fn(6);
"#);
}

// ── Static arrow function ─────────────────────────────────────

#[test] fn static_arrow_function() {
    compile_ok(r#"<?php
$double = static fn(int $x): int => $x * 2;
echo $double(21);
"#);
}

// ── Closure::bind to different class ─────────────────────────

#[test] fn closure_bind_to_different_class() {
    compile_ok(r#"<?php
class A { private string $tag = 'A'; }
class B { private string $tag = 'B'; }
$readTag = Closure::bind(function(): string { return $this->tag; }, new B(), B::class);
echo $readTag();
"#);
}

// ── Closure::fromCallable with instance method ────────────────

#[test] fn closure_from_callable_instance_method() {
    compile_ok(r#"<?php
class Formatter {
    public function format(string $s): string { return '[' . $s . ']'; }
}
$obj = new Formatter();
$fn = Closure::fromCallable([$obj, 'format']);
echo $fn('test');
"#);
}

// ── call_user_func with closure ───────────────────────────────

#[test] fn call_user_func_with_closure() {
    compile_ok(r#"<?php
$greet = function(string $name): string { return 'hi ' . $name; };
echo call_user_func($greet, 'world');
"#);
}

// ── call_user_func_array with closure ────────────────────────

#[test] fn call_user_func_array_with_closure() {
    compile_ok(r#"<?php
$sum = function(int $a, int $b, int $c): int { return $a + $b + $c; };
echo call_user_func_array($sum, [1, 2, 3]);
"#);
}

// ── Closure as stored event handler ──────────────────────────

#[test] fn closure_stored_as_event_handler() {
    compile_ok(r#"<?php
class EventEmitter {
    private array $handlers = [];
    public function on(string $event, callable $fn): void { $this->handlers[$event] = $fn; }
    public function emit(string $event, mixed $data): void {
        if (isset($this->handlers[$event])) ($this->handlers[$event])($data);
    }
}
$emitter = new EventEmitter();
$log = [];
$emitter->on('data', function(mixed $d) use (&$log): void { $log[] = $d; });
$emitter->emit('data', 'payload');
echo count($log);
"#);
}

// ── Closure with type-hinted parameter ───────────────────────

#[test] fn closure_with_typed_parameter() {
    compile_ok(r#"<?php
$stringify = function(int|float $n): string { return (string) $n; };
echo $stringify(3.14);
echo $stringify(42);
"#);
}
