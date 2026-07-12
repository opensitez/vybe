use super::helpers::{compile_ok, run_prints};

// ── Basic classes ───────────────────────────────────────────
#[test]
fn class_empty() {
    compile_ok("<?php class Foo {} $f = new Foo();");
}
#[test]
fn class_properties() {
    compile_ok(
        "<?php class Dog { public $name; public $breed; } $d = new Dog(); $d->name = 'Rex';",
    );
}
#[test]
fn class_constructor() {
    compile_ok(
        "<?php class Dog { public $name; public function __construct($name) { $this->name = $name; } } $d = new Dog('Rex'); echo $d->name;",
    );
}
#[test]
fn class_methods() {
    compile_ok(
        "<?php class Calc { public function add($a, $b) { return $a + $b; } public function sub($a, $b) { return $a - $b; } } $c = new Calc(); echo $c->add(3, 2);",
    );
}
#[test]
fn class_this() {
    compile_ok(
        "<?php class Counter { public $count = 0; public function inc() { $this->count++; return $this; } } $c = new Counter(); $c->inc()->inc();",
    );
}
#[test]
fn class_default_props() {
    compile_ok(
        "<?php class Config { public $debug = false; public $version = '1.0'; } $c = new Config(); echo $c->version;",
    );
}

// ── Inheritance ─────────────────────────────────────────────
#[test]
fn extends_basic() {
    compile_ok(
        "<?php class Animal { public $name; public function __construct($n) { $this->name = $n; } } class Dog extends Animal { public function speak() { return $this->name . ' barks'; } } $d = new Dog('Rex'); echo $d->speak();",
    );
}
#[test]
fn method_override() {
    compile_ok(
        "<?php class Base { public function greet() { return 'Hello'; } } class Child extends Base { public function greet() { return 'Hi'; } } $c = new Child(); echo $c->greet();",
    );
}
#[test]
fn parent_call() {
    compile_ok(
        "<?php class Base { public function foo() { return 'base'; } } class Child extends Base { public function foo() { return parent::foo() . '+child'; } } $c = new Child();",
    );
}

// ── Static members ──────────────────────────────────────────
#[test]
fn static_method() {
    compile_ok(
        "<?php class MathHelper { public static function square($n) { return $n * $n; } } echo MathHelper::square(5);",
    );
}
#[test]
fn class_constant() {
    compile_ok(
        "<?php class Config { const VERSION = '2.0'; const MAX = 100; } echo Config::VERSION; echo Config::MAX;",
    );
}
#[test]
fn static_factory() {
    compile_ok(
        "<?php class User { public $name; public function __construct($n) { $this->name = $n; } public static function create($n) { return new User($n); } } $u = User::create('Alice');",
    );
}

// ── Abstract classes ────────────────────────────────────────
#[test]
fn abstract_class() {
    compile_ok(
        r#"<?php
abstract class Shape {
    abstract public function area(): float;
    public function describe() { return 'Shape with area ' . $this->area(); }
}
class Circle extends Shape {
    public $radius;
    public function __construct($r) { $this->radius = $r; }
    public function area(): float { return 3.14159 * $this->radius * $this->radius; }
}
$c = new Circle(5);
echo $c->describe();
"#,
    );
}

// ── Interfaces ──────────────────────────────────────────────
#[test]
fn interface_impl() {
    compile_ok(
        r#"<?php
interface Printable { public function toString(): string; }
class Item implements Printable {
    public $name;
    public function __construct($n) { $this->name = $n; }
    public function toString(): string { return $this->name; }
}
$i = new Item('Widget');
echo $i->toString();
"#,
    );
}

// ── Traits ──────────────────────────────────────────────────
#[test]
fn trait_basic() {
    compile_ok(
        r#"<?php
trait Greetable {
    public function greet() { return 'Hello, ' . $this->name; }
}
class Person { use Greetable; public $name; public function __construct($n) { $this->name = $n; } }
$p = new Person('Alice');
echo $p->greet();
"#,
    );
}

#[test]
fn trait_multiple() {
    compile_ok(
        r#"<?php
trait HasName { public function getName() { return $this->name; } }
trait HasAge { public function getAge() { return $this->age; } }
class User {
    use HasName;
    use HasAge;
    public $name; public $age;
    public function __construct($n, $a) { $this->name = $n; $this->age = $a; }
}
$u = new User('Bob', 25);
echo $u->getName();
"#,
    );
}

// ── Enums ───────────────────────────────────────────────────
#[test]
fn enum_basic() {
    compile_ok(
        "<?php enum Color { case Red; case Green; case Blue; } $c = Color::Red; echo $c->name;",
    );
}
#[test]
fn enum_backed() {
    compile_ok(
        "<?php enum Suit: string { case Hearts = 'H'; case Diamonds = 'D'; } echo Suit::Hearts->value;",
    );
}
#[test]
fn enum_method() {
    compile_ok(
        r#"<?php
enum Status {
    case Active;
    case Inactive;
    public function label() { return $this->name; }
}
echo Status::Active->label();
"#,
    );
}

// ── Readonly / Promotion ────────────────────────────────────
#[test]
fn readonly_prop() {
    compile_ok(
        "<?php class User { public readonly string $name; public function __construct(string $n) { $this->name = $n; } } $u = new User('Alice');",
    );
}
#[test]
fn ctor_promotion() {
    compile_ok(
        "<?php class Point { public function __construct(public float $x, public float $y) {} } $p = new Point(1.0, 2.0);",
    );
}
#[test]
fn readonly_class() {
    compile_ok(
        "<?php readonly class Dto { public function __construct(public string $name, public int $age) {} } $d = new Dto('Alice', 30);",
    );
}

// ── Nullsafe ────────────────────────────────────────────────
#[test]
fn nullsafe_chain() {
    compile_ok("<?php $x = $a?->b?->c;");
}
#[test]
fn nullsafe_method() {
    compile_ok("<?php $x = $obj?->getName();");
}

// ── instanceof ──────────────────────────────────────────────
#[test]
fn instanceof_check() {
    compile_ok("<?php class A {} $a = new A(); echo $a instanceof A;");
}

// ── Fluent / chaining ───────────────────────────────────────
#[test]
fn method_chaining() {
    compile_ok(
        r#"<?php
class Builder {
    public $parts = [];
    public function add($part) { array_push($this->parts, $part); return $this; }
    public function build() { return implode(', ', $this->parts); }
}
$b = new Builder();
echo $b->add('a')->add('b')->add('c')->build();
"#,
    );
}

// ── First-class callable ────────────────────────────────────
#[test]
fn func_ref() {
    compile_ok("<?php $fn = strlen(...); echo $fn('hello');");
}
#[test]
fn method_ref() {
    compile_ok(
        "<?php class A { public function foo() { return 42; } } $a = new A(); $fn = $a->foo(...);",
    );
}
#[test]
fn static_ref() {
    compile_ok(
        "<?php class M { public static function sq($n) { return $n * $n; } } $fn = M::sq(...);",
    );
}

#[test]
fn variadic_static_method_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Greeter {
    public static function call($prefix, ...$parts) {
        echo $prefix . ':' . implode(',', $parts);
    }
}
Greeter::call('head', 'a', 'b', 'c');
"#
        ),
        &["head:a,b,c"]
    );
}
