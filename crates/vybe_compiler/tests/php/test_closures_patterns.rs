use super::helpers::run_prints;

// ── Closure::bind and bindTo ──────────────────────────────────

#[test] fn closure_bind_to_object() {
    assert_eq!(run_prints(r#"<?php
class Obj { private int $x = 42; }
$fn = Closure::bind(function() { return $this->x; }, new Obj, Obj::class);
echo $fn();
"#), vec!["42"]);
}
#[test] fn closure_bind_to_static_context() {
    assert_eq!(run_prints(r#"<?php
class MyClass { private static string $secret = 'hidden'; }
$fn = Closure::bind(function() { return static::$secret; }, null, MyClass::class);
echo $fn();
"#), vec!["hidden"]);
}
#[test] fn closure_bind_to_child_accesses_parent_private() {
    assert_eq!(run_prints(r#"<?php
class Base { private int $val = 99; }
class Derived extends Base {}
$fn = Closure::bind(function() { return $this->val; }, new Derived, Base::class);
echo $fn();
"#), vec!["99"]);
}

// ── Closure::fromCallable ─────────────────────────────────────

#[test] fn closure_from_callable_function() {
    assert_eq!(run_prints(r#"<?php
$fn = Closure::fromCallable('strtoupper');
echo $fn('hello');
"#), vec!["HELLO"]);
}
#[test] fn closure_from_callable_method() {
    assert_eq!(run_prints(r#"<?php
class Formatter { public function upper(string $s): string { return strtoupper($s); } }
$f = new Formatter;
$fn = Closure::fromCallable([$f, 'upper']);
echo $fn('world');
"#), vec!["WORLD"]);
}
#[test] fn closure_from_callable_static_method() {
    assert_eq!(run_prints(r#"<?php
class Math { public static function square(int $n): int { return $n * $n; } }
$fn = Closure::fromCallable(['Math', 'square']);
echo $fn(7);
"#), vec!["49"]);
}

// ── Static closures ───────────────────────────────────────────

#[test] fn static_closure_no_this() {
    assert_eq!(run_prints(r#"<?php
$fn = static fn($x) => $x * 2;
echo $fn(5);
"#), vec!["10"]);
}
#[test] fn static_closure_cannot_bind_this() {
    assert_eq!(run_prints(r#"<?php
class Obj { private int $v = 1; }
$fn = static function() { return 42; };
$bound = Closure::bind($fn, new Obj, Obj::class);
echo $bound();
"#), vec!["42"]);
}

// ── Recursive closures ────────────────────────────────────────

#[test] fn recursive_closure_self_reference() {
    assert_eq!(run_prints(r#"<?php
$factorial = null;
$factorial = function(int $n) use (&$factorial): int {
    return $n <= 1 ? 1 : $n * $factorial($n - 1);
};
echo $factorial(6);
"#), vec!["720"]);
}
#[test] fn recursive_closure_fibonacci() {
    assert_eq!(run_prints(r#"<?php
$fib = null;
$fib = function(int $n) use (&$fib): int {
    if ($n <= 1) return $n;
    return $fib($n-1) + $fib($n-2);
};
echo $fib(10);
"#), vec!["55"]);
}

// ── Closures as event handlers ────────────────────────────────

#[test] fn closure_event_handler_pattern() {
    assert_eq!(run_prints(r#"<?php
class Button {
    private array $handlers = [];
    public function onClick(Closure $fn): void { $this->handlers[] = $fn; }
    public function click(string $data): void { foreach ($this->handlers as $h) $h($data); }
}
$log = [];
$btn = new Button;
$btn->onClick(function($d) use (&$log) { $log[] = "clicked:$d"; });
$btn->onClick(function($d) use (&$log) { $log[] = "handled:$d"; });
$btn->click('left');
echo implode(',', $log);
"#), vec!["clicked:left,handled:left"]);
}

// ── Arrow functions ───────────────────────────────────────────

#[test] fn arrow_fn_captures_outer_scope() {
    assert_eq!(run_prints(r#"<?php
$multiplier = 3;
$fn = fn($n) => $n * $multiplier;
echo $fn(7);
"#), vec!["21"]);
}
#[test] fn arrow_fn_nested_capture() {
    assert_eq!(run_prints(r#"<?php
$a = 1; $b = 2; $c = 3;
$fn = fn($x) => fn($y) => $x + $y + $a + $b + $c;
echo $fn(10)(20);
"#), vec!["36"]);
}
#[test] fn arrow_fn_in_array_map() {
    assert_eq!(run_prints(r#"<?php
$base = 100;
$result = array_map(fn($n) => $base + $n, [1, 2, 3, 4, 5]);
echo implode(',', $result);
"#), vec!["101,102,103,104,105"]);
}
#[test] fn arrow_fn_with_type_hints() {
    assert_eq!(run_prints(r#"<?php
$double = fn(int $n): int => $n * 2;
echo $double(21);
"#), vec!["42"]);
}

// ── Closure scope isolation ───────────────────────────────────

#[test] fn closure_captures_value_not_binding() {
    assert_eq!(run_prints(r#"<?php
$x = 5;
$fn = function() use ($x) { return $x; };
$x = 999;
echo $fn();
"#), vec!["5"]);
}
#[test] fn closure_captures_reference_sees_change() {
    assert_eq!(run_prints(r#"<?php
$x = 5;
$fn = function() use (&$x) { return $x; };
$x = 999;
echo $fn();
"#), vec!["999"]);
}
