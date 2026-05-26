use super::helpers::run_prints;

// ── array_is_list (PHP 8.1) ───────────────────────────────────

#[test] fn array_is_list_sequential() {
    assert_eq!(run_prints(r#"<?php echo array_is_list([1,2,3]) ? 'yes' : 'no'; "#), vec!["yes"]);
}
#[test] fn array_is_list_empty() {
    assert_eq!(run_prints(r#"<?php echo array_is_list([]) ? 'yes' : 'no'; "#), vec!["yes"]);
}
#[test] fn array_is_list_string_keys() {
    assert_eq!(run_prints(r#"<?php echo array_is_list(['a'=>1,'b'=>2]) ? 'yes' : 'no'; "#), vec!["no"]);
}
#[test] fn array_is_list_gap_in_keys() {
    assert_eq!(run_prints(r#"<?php $a = [0=>1,2=>3]; echo array_is_list($a) ? 'yes' : 'no'; "#), vec!["no"]);
}
#[test] fn array_is_list_after_unset() {
    assert_eq!(run_prints(r#"<?php $a = [1,2,3]; unset($a[1]); echo array_is_list($a) ? 'yes' : 'no'; "#), vec!["no"]);
}

// ── Readonly properties (PHP 8.1) ─────────────────────────────

#[test] fn readonly_property_set_once() {
    assert_eq!(run_prints(r#"<?php
class User { public function __construct(public readonly string $name) {} }
$u = new User('Alice');
echo $u->name;
"#), vec!["Alice"]);
}
#[test] fn readonly_property_throws_on_reassign() {
    assert_eq!(run_prints(r#"<?php
class User { public function __construct(public readonly string $name) {} }
$u = new User('Alice');
try { $u->name = 'Bob'; } catch (Error $e) { echo 'readonly'; }
"#), vec!["readonly"]);
}
#[test] fn readonly_property_in_clone() {
    assert_eq!(run_prints(r#"<?php
class VO { public function __construct(public readonly int $val) {} }
$a = new VO(1);
$b = clone $a;
echo $a->val . ',' . $b->val;
"#), vec!["1,1"]);
}

// ── Enums (PHP 8.1) ───────────────────────────────────────────

#[test] fn enum_pure_unit_enum() {
    assert_eq!(run_prints(r#"<?php
enum Direction { case North; case South; case East; case West; }
$d = Direction::North;
echo $d->name;
"#), vec!["North"]);
}
#[test] fn enum_backed_string() {
    assert_eq!(run_prints(r#"<?php
enum Color: string { case Red = 'red'; case Blue = 'blue'; }
echo Color::Blue->value . ',' . Color::from('red')->name;
"#), vec!["blue,Red"]);
}
#[test] fn enum_backed_int() {
    assert_eq!(run_prints(r#"<?php
enum Priority: int { case Low = 1; case Mid = 2; case High = 3; }
$p = Priority::High;
echo $p->value;
"#), vec!["3"]);
}
#[test] fn enum_implements_interface() {
    assert_eq!(run_prints(r#"<?php
interface HasLabel { public function label(): string; }
enum Status: string implements HasLabel {
    case Active = 'active';
    case Inactive = 'inactive';
    public function label(): string { return ucfirst($this->value); }
}
echo Status::Active->label();
"#), vec!["Active"]);
}
#[test] fn enum_in_match() {
    assert_eq!(run_prints(r#"<?php
enum Suit { case Hearts; case Diamonds; case Clubs; case Spades; }
$s = Suit::Hearts;
echo match($s) {
    Suit::Hearts, Suit::Diamonds => 'red',
    default => 'black',
};
"#), vec!["red"]);
}

// ── Intersection types (PHP 8.1) ──────────────────────────────

#[test] fn intersection_type_accepted() {
    assert_eq!(run_prints(r#"<?php
interface Stringable2 { public function __toString(): string; }
interface Serializable2 { public function serialize(): string; }
class Item implements Stringable2, Serializable2 {
    public function __toString(): string { return 'item'; }
    public function serialize(): string { return 'serialized'; }
}
function process(Stringable2&Serializable2 $obj): string {
    return (string)$obj . ':' . $obj->serialize();
}
echo process(new Item);
"#), vec!["item:serialized"]);
}

// ── Fibers (PHP 8.1) ──────────────────────────────────────────

#[test] fn fiber_basic_suspend_resume() {
    assert_eq!(run_prints(r#"<?php
$f = new Fiber(function(): void {
    echo 'A';
    Fiber::suspend();
    echo 'C';
});
echo 'start,';
$f->start();
echo 'B,';
$f->resume();
echo 'done';
"#), vec!["start,AB,Cdone"]);
}
#[test] fn fiber_passes_value_on_suspend() {
    assert_eq!(run_prints(r#"<?php
$f = new Fiber(function(): string {
    $v = Fiber::suspend('mid');
    return 'end:' . $v;
});
$mid = $f->start();
echo $mid . ',';
$f->resume('resumed');
echo $f->getReturn();
"#), vec!["mid,end:resumed"]);
}

// ── never return type (PHP 8.1) ───────────────────────────────

#[test] fn never_function_throws() {
    assert_eq!(run_prints(r#"<?php
function abort(string $msg): never { throw new RuntimeException($msg); }
try { abort('fatal'); } catch (RuntimeException $e) { echo $e->getMessage(); }
"#), vec!["fatal"]);
}

// ── First-class callables (PHP 8.1) ───────────────────────────

#[test] fn first_class_callable_builtin() {
    assert_eq!(run_prints(r#"<?php
$fn = strlen(...);
echo $fn('hello');
"#), vec!["5"]);
}
#[test] fn first_class_callable_method() {
    assert_eq!(run_prints(r#"<?php
class Math { public function double(int $n): int { return $n * 2; } }
$m = new Math;
$fn = $m->double(...);
echo $fn(6);
"#), vec!["12"]);
}
#[test] fn first_class_callable_in_array_map() {
    assert_eq!(run_prints(r#"<?php
echo implode(',', array_map(strtoupper(...), ['a','b','c']));
"#), vec!["A,B,C"]);
}
