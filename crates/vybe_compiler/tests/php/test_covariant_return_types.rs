use super::helpers::{compile_ok, run_prints};

// ── Covariant return types — child returns narrower type ──────

#[test] fn child_overrides_with_own_class_return_type() {
    assert_eq!(run_prints(r#"<?php
class Animal { public function create(): static { return new static(); } }
class Dog extends Animal {}
$dog = Dog::create();
echo get_class($dog);
"#), vec!["Dog"]);
}

#[test] fn covariant_return_self_in_fluent_builder() {
    assert_eq!(run_prints(r#"<?php
class Builder {
    protected array $parts = [];
    public function add(string $part): static {
        $this->parts[] = $part;
        return $this;
    }
    public function build(): string { return implode(',', $this->parts); }
}
class FancyBuilder extends Builder {
    public function addFancy(string $part): static {
        return $this->add("*$part*");
    }
}
echo (new FancyBuilder())->addFancy('a')->add('b')->build();
"#), vec!["*a*,b"]);
}

#[test] fn covariant_return_child_class_instead_of_parent() {
    assert_eq!(run_prints(r#"<?php
class Shape { public function describe(): string { return "shape"; } }
class Circle extends Shape { public function describe(): string { return "circle"; } }
class ShapeFactory {
    public function make(): Shape { return new Shape(); }
}
class CircleFactory extends ShapeFactory {
    public function make(): Circle { return new Circle(); }
}
$factory = new CircleFactory();
echo $factory->make()->describe();
"#), vec!["circle"]);
}

#[test] fn covariant_return_interface_to_concrete_class() {
    assert_eq!(run_prints(r#"<?php
interface Cloneable2 { public function clone(): static; }
class Point implements Cloneable2 {
    public function __construct(public int $x, public int $y) {}
    public function clone(): static { return new static($this->x, $this->y); }
}
$p = new Point(3, 4);
$q = $p->clone();
echo $q->x . ',' . $q->y;
"#), vec!["3,4"]);
}

// ── static return type ────────────────────────────────────────

#[test] fn static_return_type_returns_correct_class() {
    assert_eq!(run_prints(r#"<?php
class Registry {
    private static array $items = [];
    public static function add(string $item): static {
        static::$items[] = $item;
        return new static();
    }
    public static function count(): int { return count(static::$items); }
}
Registry::add('a');
Registry::add('b');
echo Registry::count();
"#), vec!["2"]);
}

#[test] fn static_return_named_constructor() {
    assert_eq!(run_prints(r#"<?php
class Money {
    private function __construct(private int $cents) {}
    public static function fromCents(int $cents): static { return new static($cents); }
    public function amount(): int { return $this->cents; }
}
class Euro extends Money {}
$e = Euro::fromCents(500);
echo $e->amount();
"#), vec!["500"]);
}

// ── Contravariant parameter types ────────────────────────────

#[test] fn contravariant_parameter_widens_type() {
    assert_eq!(run_prints(r#"<?php
class Animal { public function name(): string { return "animal"; } }
class Dog extends Animal { public function name(): string { return "dog"; } }
interface Feeder { public function feed(Dog $dog): void; }
class GenericFeeder implements Feeder {
    public function feed(Animal $animal): void { echo "feeding " . $animal->name(); }
}
$feeder = new GenericFeeder();
$feeder->feed(new Dog());
"#), vec!["feeding dog"]);
}

// ── never return type ─────────────────────────────────────────

#[test] fn never_return_type_function_throws() {
    assert_eq!(run_prints(r#"<?php
function fail(string $msg): never {
    throw new RuntimeException($msg);
}
try {
    fail("oops");
} catch (RuntimeException $e) {
    echo $e->getMessage();
}
"#), vec!["oops"]);
}

#[test] fn never_return_type_function_exits() {
    compile_ok(r#"<?php
function abort(int $code): never {
    throw new RuntimeException("abort: $code");
}
"#);
}

// ── void return type ─────────────────────────────────────────

#[test] fn void_return_type_function_returns_nothing() {
    assert_eq!(run_prints(r#"<?php
function printLine(string $s): void { echo $s; }
printLine("hello");
"#), vec!["hello"]);
}

#[test] fn void_return_implicit_null() {
    assert_eq!(run_prints(r#"<?php
function doNothing(): void {}
$result = doNothing();
echo var_export($result, true);
"#), vec!["NULL"]);
}

// ── mixed return type ─────────────────────────────────────────

#[test] fn mixed_return_type_accepts_any() {
    assert_eq!(run_prints(r#"<?php
function identity(mixed $v): mixed { return $v; }
echo identity(42) . ',' . identity("hello");
"#), vec!["42,hello"]);
}

// ── Union return types ────────────────────────────────────────

#[test] fn union_return_type_int_or_false() {
    assert_eq!(run_prints(r#"<?php
function search(array $arr, int $val): int|false {
    $idx = array_search($val, $arr);
    return $idx !== false ? $idx : false;
}
echo search([10, 20, 30], 20);
echo ',';
echo var_export(search([10, 20, 30], 99), true);
"#), vec!["1,false"]);
}

#[test] fn union_return_type_string_or_null() {
    assert_eq!(run_prints(r#"<?php
function findName(array $map, int $id): string|null {
    return $map[$id] ?? null;
}
echo findName([1 => 'Alice', 2 => 'Bob'], 1);
echo ',';
echo var_export(findName([1 => 'Alice'], 99), true);
"#), vec!["Alice,NULL"]);
}

// ── Nullable return type ──────────────────────────────────────

#[test] fn nullable_return_type_returns_null() {
    assert_eq!(run_prints(r#"<?php
function maybeValue(bool $flag): ?string {
    return $flag ? "yes" : null;
}
echo maybeValue(true) . ',' . var_export(maybeValue(false), true);
"#), vec!["yes,NULL"]);
}

// ── Return type in interface + implementation ─────────────────

#[test] fn interface_return_type_enforced_in_impl() {
    assert_eq!(run_prints(r#"<?php
interface Transformer { public function transform(string $s): string; }
class UpperTransformer implements Transformer {
    public function transform(string $s): string { return strtoupper($s); }
}
$t = new UpperTransformer();
echo $t->transform("hello");
"#), vec!["HELLO"]);
}

// ── Abstract method return type enforced ──────────────────────

#[test] fn abstract_return_type_enforced_in_child() {
    assert_eq!(run_prints(r#"<?php
abstract class Serializer {
    abstract public function serialize(array $data): string;
}
class JsonSerializer extends Serializer {
    public function serialize(array $data): string { return json_encode($data); }
}
$s = new JsonSerializer();
echo $s->serialize(['a' => 1]);
"#), vec!["{\"a\":1}"]);
}

// ── self vs static in return type ────────────────────────────

#[test] fn self_return_type_stays_in_declared_class() {
    assert_eq!(run_prints(r#"<?php
class Base {
    public function withData(string $d): self {
        $clone = clone $this;
        return $clone;
    }
    public function type(): string { return "Base"; }
}
$b = new Base();
echo $b->withData("x")->type();
"#), vec!["Base"]);
}

#[test] fn static_return_type_resolves_to_subclass() {
    assert_eq!(run_prints(r#"<?php
class Node {
    public function next(): static { return new static(); }
    public function className(): string { return static::class; }
}
class ListNode extends Node {}
echo (new ListNode())->next()->className();
"#), vec!["ListNode"]);
}

// ── Intersection type as return ───────────────────────────────

#[test] fn intersection_return_type_both_interfaces_satisfied() {
    assert_eq!(run_prints(r#"<?php
interface Countable2 { public function count(): int; }
interface Listable { public function toList(): array; }
class Collection implements Countable2, Listable {
    private array $items;
    public function __construct(array $items) { $this->items = $items; }
    public function count(): int { return count($this->items); }
    public function toList(): array { return $this->items; }
}
function getCollection(): Countable2&Listable { return new Collection([1,2,3]); }
$c = getCollection();
echo $c->count();
"#), vec!["3"]);
}

// ── Parent return type widened correctly ──────────────────────

#[test] fn parent_method_return_type_usable_when_overridden() {
    assert_eq!(run_prints(r#"<?php
class A { public function value(): int { return 1; } }
class B extends A { public function value(): int { return parent::value() + 10; } }
echo (new B())->value();
"#), vec!["11"]);
}

// ── Chained covariant returns ─────────────────────────────────

#[test] fn chained_covariant_static_returns() {
    assert_eq!(run_prints(r#"<?php
class Query {
    protected array $conditions = [];
    public function where(string $cond): static { $this->conditions[] = $cond; return $this; }
    public function sql(): string { return implode(' AND ', $this->conditions); }
}
class UserQuery extends Query {
    public function active(): static { return $this->where('active = 1'); }
}
echo (new UserQuery())->active()->where('age > 18')->sql();
"#), vec!["active = 1 AND age > 18"]);
}
