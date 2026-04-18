use super::helpers;
use helpers::compile_ok;

fn parse_ok(src: &str) -> bool {
    vybec::parser_php::Parser::new(src).and_then(|mut p| p.parse_program()).is_ok()
}

// ── Complex string interpolation ────────────────────────────
#[test] fn interp_array_index() {
    compile_ok(r#"<?php $arr = [1,2,3]; echo "first: {$arr[0]}";"#);
}

#[test] fn interp_object_prop() {
    compile_ok(r#"<?php $obj = new stdClass(); echo "name: {$obj->name}";"#);
}

#[test] fn interp_simple_array() {
    // $arr[0] without braces — direct interpolation
    compile_ok(r#"<?php $arr = ['a','b']; echo "val: $arr[0]";"#);
}

#[test] fn interp_simple_prop() {
    // $obj->prop without braces — direct interpolation
    compile_ok(r#"<?php $o = new stdClass(); echo "val: $o->name";"#);
}

// ── Enum ->value / ->name ───────────────────────────────────
#[test] fn enum_case_access() {
    compile_ok(r#"<?php
enum Color { case Red; case Green; case Blue; }
$c = Color::Red;
echo $c->name;
echo $c->value;
"#);
}

#[test] fn enum_backed_access() {
    compile_ok(r#"<?php
enum Suit: string {
    case Hearts = 'H';
    case Diamonds = 'D';
    case Clubs = 'C';
    case Spades = 'S';
}
$s = Suit::Hearts;
echo $s->value;
echo $s->name;
"#);
}

// ── Match as statement ──────────────────────────────────────
#[test] fn match_statement() {
    compile_ok(r#"<?php
$x = 2;
match($x) {
    1 => 'one',
    2 => 'two',
    default => 'other'
};
"#);
}

#[test] fn match_with_calls() {
    compile_ok(r#"<?php
$action = 'greet';
match($action) {
    'greet' => 'hello',
    'bye' => 'goodbye',
    default => 'unknown'
};
"#);
}

// ── Fiber multi-arg start ───────────────────────────────────
#[test] fn fiber_multi_arg_start() {
    compile_ok(r#"<?php
$fiber = new Fiber(function($a, $b, $c) {
    return $a + $b + $c;
});
$result = $fiber->start(10, 20, 30);
"#);
}

#[test] fn fiber_single_arg_start() {
    compile_ok(r#"<?php
$fiber = new Fiber(function($x) {
    return $x * 2;
});
$result = $fiber->start(21);
"#);
}

#[test] fn fiber_no_arg_start() {
    compile_ok(r#"<?php
$fiber = new Fiber(function() {
    return 42;
});
$result = $fiber->start();
"#);
}

// ── Intersection types ──────────────────────────────────────
#[test] fn intersection_type_parse() {
    assert!(parse_ok("<?php function foo(A&B $x): void {}"), "intersection type parse");
}

#[test] fn intersection_return_type() {
    assert!(parse_ok("<?php function bar(): Countable&Iterator { return null; }"), "intersection return type");
}

// ── Arrow function auto-capture ─────────────────────────────
#[test] fn arrow_fn_auto_capture() {
    compile_ok(r#"<?php
$multiplier = 3;
$fn = fn($x) => $x * $multiplier;
echo $fn(5);
"#);
}

#[test] fn arrow_fn_capture_multiple() {
    compile_ok(r#"<?php
$base = 100;
$tax = 0.2;
$calc = fn($price) => ($price + $base) * (1 + $tax);
echo $calc(50);
"#);
}

#[test] fn arrow_fn_nested_capture() {
    compile_ok(r#"<?php
$x = 10;
$outer = fn($a) => fn($b) => $a + $b + $x;
"#);
}

// ── Attributes (parse only) ─────────────────────────────────
#[test] fn attribute_on_function() {
    assert!(parse_ok("<?php #[Pure] function add(int $a, int $b): int { return $a + $b; }"));
}

#[test] fn attribute_on_class() {
    assert!(parse_ok("<?php #[Entity] #[Table('users')] class User {}"));
}

#[test] fn attribute_with_args() {
    assert!(parse_ok("<?php #[Route('/api', methods: ['GET'])] function handler() {}"));
}

// ── Readonly properties ─────────────────────────────────────
#[test] fn readonly_property_class() {
    compile_ok(r#"<?php
class User {
    public readonly string $name;
    public function __construct(string $name) {
        $this->name = $name;
    }
}
$u = new User('Alice');
echo $u->name;
"#);
}

// ── Constructor promotion ───────────────────────────────────
#[test] fn ctor_promotion_full() {
    compile_ok(r#"<?php
class Point {
    public function __construct(
        public float $x,
        public float $y,
        public float $z = 0.0
    ) {}
}
$p = new Point(1.0, 2.0);
"#);
}

// ── Const type hints ────────────────────────────────────────
#[test] fn class_const_typed() {
    compile_ok("<?php class A { const VERSION = '1.0'; const MAX = 100; }");
}

// ── Abstract with interface implementation ───────────────────
#[test] fn abstract_implements() {
    compile_ok(r#"<?php
abstract class Vehicle {
    abstract public function wheels(): int;
    public function describe() { return 'Vehicle with ' . $this->wheels() . ' wheels'; }
}
class Car extends Vehicle {
    public function wheels(): int { return 4; }
}
$car = new Car();
"#);
}
