use super::helpers::run_prints;

// ── Late static binding ──────────────────────────────────────────
#[test]
fn late_static_binding_basic() {
    assert_eq!(run_prints(r#"<?php
class Base {
    protected static string $type = "base";
    public static function getType(): string {
        return static::$type;
    }
}
class Child extends Base {
    protected static string $type = "child";
}
echo Base::getType();
echo Child::getType();
"#), &["base", "child"]);
}

#[test]
fn late_static_binding_new_static() {
    assert_eq!(run_prints(r#"<?php
class Animal {
    public string $name;
    public function __construct(string $name) {
        $this->name = $name;
    }
    public static function create(string $name): static {
        return new static($name);
    }
    public function type(): string { return "animal"; }
}
class Dog extends Animal {
    public function type(): string { return "dog"; }
}
$a = Animal::create("Rex");
$d = Dog::create("Buddy");
echo $a->type();
echo $d->type();
echo $d->name;
"#), &["animal", "dog", "Buddy"]);
}

#[test]
fn late_static_binding_constant() {
    assert_eq!(run_prints(r#"<?php
class Shape {
    const NAME = "shape";
    public static function describe(): string {
        return static::NAME;
    }
}
class Circle extends Shape {
    const NAME = "circle";
}
echo Shape::describe();
echo Circle::describe();
"#), &["shape", "circle"]);
}

// ── Anonymous classes ────────────────────────────────────────────
#[test]
fn anonymous_class_basic() {
    assert_eq!(run_prints(r#"<?php
$obj = new class {
    public function greet(): string {
        return "hello from anonymous";
    }
};
echo $obj->greet();
"#), &["hello from anonymous"]);
}

#[test]
fn anonymous_class_with_constructor() {
    assert_eq!(run_prints(r#"<?php
$obj = new class("world") {
    public function __construct(private string $name) {}
    public function greet(): string {
        return "hello $this->name";
    }
};
echo $obj->greet();
"#), &["hello world"]);
}

#[test]
fn anonymous_class_implements_interface() {
    assert_eq!(run_prints(r#"<?php
interface Printable {
    public function toString(): string;
}
$obj = new class implements Printable {
    public function toString(): string {
        return "I am printable";
    }
};
echo $obj->toString();
"#), &["I am printable"]);
}

#[test]
fn anonymous_class_extends() {
    assert_eq!(run_prints(r#"<?php
class Logger {
    public function log(string $msg): void {
        echo "LOG: $msg";
    }
}
$obj = new class extends Logger {
    public function info(string $msg): void {
        $this->log("INFO: $msg");
    }
};
$obj->info("started");
"#), &["LOG: INFO: started"]);
}

// ── Trait conflict resolution ────────────────────────────────────
#[test]
fn trait_insteadof() {
    assert_eq!(run_prints(r#"<?php
trait A {
    public function hello(): string { return "A"; }
}
trait B {
    public function hello(): string { return "B"; }
}
class C {
    use A, B {
        A::hello insteadof B;
    }
}
$c = new C();
echo $c->hello();
"#), &["A"]);
}

#[test]
fn trait_alias_as() {
    assert_eq!(run_prints(r#"<?php
trait Greeter {
    public function hello(): string { return "hello"; }
}
class MyClass {
    use Greeter {
        hello as greet;
    }
}
$obj = new MyClass();
echo $obj->greet();
echo $obj->hello();
"#), &["hello", "hello"]);
}

#[test]
fn trait_insteadof_with_alias() {
    assert_eq!(run_prints(r#"<?php
trait X {
    public function speak(): string { return "X speaks"; }
}
trait Y {
    public function speak(): string { return "Y speaks"; }
}
class Z {
    use X, Y {
        X::speak insteadof Y;
        Y::speak as ySpeak;
    }
}
$z = new Z();
echo $z->speak();
echo $z->ySpeak();
"#), &["X speaks", "Y speaks"]);
}

// ── Abstract methods in traits ───────────────────────────────────
#[test]
fn trait_abstract_method() {
    assert_eq!(run_prints(r#"<?php
trait Validatable {
    abstract protected function validate(): bool;
    public function isValid(): string {
        return $this->validate() ? "valid" : "invalid";
    }
}
class Email {
    use Validatable;
    public function __construct(private string $value) {}
    protected function validate(): bool {
        return str_contains($this->value, "@");
    }
}
$e1 = new Email("user@example.com");
$e2 = new Email("invalid");
echo $e1->isValid();
echo $e2->isValid();
"#), &["valid", "invalid"]);
}

// ── Trait with constants (PHP 8.2) ───────────────────────────────
#[test]
fn trait_with_constants() {
    assert_eq!(run_prints(r#"<?php
trait HasVersion {
    const VERSION = "1.0";
    public function getVersion(): string {
        return self::VERSION;
    }
}
class App {
    use HasVersion;
}
$app = new App();
echo $app->getVersion();
"#), &["1.0"]);
}

// ── Interface constants and default methods ──────────────────────
#[test]
fn interface_with_constants() {
    assert_eq!(run_prints(r#"<?php
interface Status {
    const ACTIVE = 1;
    const INACTIVE = 0;
}
class User implements Status {
    public function getStatus(): int {
        return self::ACTIVE;
    }
}
$u = new User();
echo $u->getStatus();
echo User::INACTIVE;
"#), &["1", "0"]);
}

// ── Multiple interface implementation ────────────────────────────
#[test]
fn multiple_interfaces() {
    assert_eq!(run_prints(r#"<?php
interface Readable {
    public function read(): string;
}
interface Writable {
    public function write(string $data): void;
}
class File implements Readable, Writable {
    private string $content = "";
    public function read(): string { return $this->content; }
    public function write(string $data): void { $this->content .= $data; }
}
$f = new File();
$f->write("hello");
$f->write(" world");
echo $f->read();
"#), &["hello world"]);
}

// ── Abstract class with concrete methods ─────────────────────────
#[test]
fn abstract_template_method() {
    assert_eq!(run_prints(r#"<?php
abstract class Report {
    abstract protected function getData(): array;
    public function generate(): string {
        $data = $this->getData();
        return implode(", ", $data);
    }
}
class SalesReport extends Report {
    protected function getData(): array {
        return ["Q1: 100", "Q2: 200", "Q3: 150"];
    }
}
$r = new SalesReport();
echo $r->generate();
"#), &["Q1: 100, Q2: 200, Q3: 150"]);
}

// ── Covariant return types ───────────────────────────────────────
#[test]
fn covariant_return() {
    assert_eq!(run_prints(r#"<?php
class Collection {
    protected array $items;
    public function __construct(array $items) { $this->items = $items; }
    public function first(): mixed { return $this->items[0] ?? null; }
}
class TypedCollection extends Collection {
    public function first(): string {
        return (string) parent::first();
    }
}
$c = new TypedCollection(["hello", "world"]);
echo $c->first();
"#), &["hello"]);
}

// ── Property hooks / accessors pattern ───────────────────────────
#[test]
fn getter_setter_pattern() {
    assert_eq!(run_prints(r#"<?php
class Temperature {
    private float $celsius;
    public function __construct(float $celsius) {
        $this->setCelsius($celsius);
    }
    public function getCelsius(): float { return $this->celsius; }
    public function setCelsius(float $val): void { $this->celsius = $val; }
    public function getFahrenheit(): float {
        return $this->celsius * 9 / 5 + 32;
    }
}
$t = new Temperature(100);
echo $t->getCelsius();
echo $t->getFahrenheit();
"#), &["100", "212"]);
}

// ── Enum implementing interface ──────────────────────────────────
#[test]
fn enum_implements_interface() {
    assert_eq!(run_prints(r#"<?php
interface HasLabel {
    public function label(): string;
}
enum Color: string implements HasLabel {
    case Red = "red";
    case Green = "green";
    case Blue = "blue";
    public function label(): string {
        return strtoupper($this->value);
    }
}
echo Color::Red->label();
echo Color::Green->value;
"#), &["RED", "green"]);
}

// ── Enum from / tryFrom ──────────────────────────────────────────
#[test]
fn enum_from_tryfrom() {
    assert_eq!(run_prints(r#"<?php
enum Suit: string {
    case Hearts = "H";
    case Diamonds = "D";
    case Clubs = "C";
    case Spades = "S";
}
$s = Suit::from("H");
echo $s->name;
$t = Suit::tryFrom("X");
echo $t === null ? "null" : $t->name;
"#), &["Hearts", "null"]);
}

// ── Readonly constructor promotion ───────────────────────────────
#[test]
fn readonly_promotion_combo() {
    assert_eq!(run_prints(r#"<?php
class Point {
    public function __construct(
        public readonly float $x,
        public readonly float $y,
        public readonly float $z = 0.0,
    ) {}
    public function distanceTo(Point $other): float {
        return sqrt(
            ($this->x - $other->x) ** 2 +
            ($this->y - $other->y) ** 2 +
            ($this->z - $other->z) ** 2
        );
    }
}
$a = new Point(0, 0, 0);
$b = new Point(3, 4, 0);
echo $a->distanceTo($b);
"#), &["5"]);
}

// ── Intersection types ───────────────────────────────────────────
#[test]
fn intersection_type_param() {
    assert_eq!(run_prints(r#"<?php
interface Countable2 {
    public function count(): int;
}
interface Stringable2 {
    public function __toString(): string;
}
class Items implements Countable2, Stringable2 {
    private array $data;
    public function __construct(array $data) { $this->data = $data; }
    public function count(): int { return count($this->data); }
    public function __toString(): string { return implode(",", $this->data); }
}
function describe(Countable2&Stringable2 $obj): void {
    echo $obj->count();
    echo $obj;
}
describe(new Items(["a", "b", "c"]));
"#), &["3", "a,b,c"]);
}

// ── Named arguments ──────────────────────────────────────────────
#[test]
fn named_args_skip_defaults() {
    assert_eq!(run_prints(r#"<?php
function createTag(string $tag, string $content, string $class = "", string $id = ""): string {
    $attrs = "";
    if ($class) $attrs .= " class=\"$class\"";
    if ($id) $attrs .= " id=\"$id\"";
    return "<$tag$attrs>$content</$tag>";
}
echo createTag("div", "hello", id: "main");
echo createTag(tag: "span", content: "world", class: "bold");
"#), &["<div id=\"main\">hello</div>", "<span class=\"bold\">world</span>"]);
}

// ── Stringable interface ─────────────────────────────────────────
#[test]
fn stringable_interface() {
    assert_eq!(run_prints(r#"<?php
class Money implements Stringable {
    public function __construct(private int $cents) {}
    public function __toString(): string {
        return "$" . number_format($this->cents / 100, 2);
    }
}
function display(Stringable $item): void {
    echo $item;
}
display(new Money(1299));
display(new Money(50));
"#), &["$12.99", "$0.50"]);
}
