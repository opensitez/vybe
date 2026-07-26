use super::helpers::{compile_ok, run_prints};

// ── Late static binding ──────────────────────────────────────────
#[test]
fn late_static_binding_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    protected static string $type = "base";
    public static function getType(): string {
        return static::$type;
    }
}
class Child extends Base {
    protected static string $type = "child";
}
echo Base::getType(), "\n";
echo Child::getType(), "\n";
"#
        ),
        &["base", "child"]
    );
}

#[test]
fn late_static_binding_new_static() {
    assert_eq!(
        run_prints(
            r#"<?php
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
echo $a->type(), "\n";
echo $d->type(), "\n";
echo $d->name, "\n";
"#
        ),
        &["animal", "dog", "Buddy"]
    );
}

#[test]
fn late_static_binding_constant() {
    assert_eq!(
        run_prints(
            r#"<?php
class Shape {
    const NAME = "shape";
    public static function describe(): string {
        return static::NAME;
    }
}
class Circle extends Shape {
    const NAME = "circle";
}
echo Shape::describe(), "\n";
echo Circle::describe(), "\n";
"#
        ),
        &["shape", "circle"]
    );
}

#[test]
fn late_static_binding_self_vs_static() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public static function selfClass(): string { return self::class; }
    public static function staticClass(): string { return static::class; }
}
class Child extends Base {}
echo Base::selfClass(), "\n";
echo Child::selfClass(), "\n";
echo Base::staticClass(), "\n";
echo Child::staticClass(), "\n";
"#
        ),
        &["Base", "Base", "Base", "Child"]
    );
}

#[test]
fn late_static_binding_factory_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
class Vehicle {
    protected string $color = "white";
    public static function make(): static {
        return new static();
    }
    public function paint(string $c): static {
        $clone = clone $this;
        $clone->color = $c;
        return $clone;
    }
    public function describe(): string {
        return static::class . ":" . $this->color;
    }
}
class Car extends Vehicle {}
$v = Vehicle::make()->paint("red");
$c = Car::make()->paint("blue");
echo $v->describe(), "\n";
echo $c->describe(), "\n";
"#
        ),
        &["Vehicle:red", "Car:blue"]
    );
}

// ── Anonymous classes ────────────────────────────────────────────
#[test]
fn anonymous_class_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new class {
    public function greet(): string {
        return "hello from anonymous";
    }
};
echo $obj->greet(), "\n";
"#
        ),
        &["hello from anonymous"]
    );
}

#[test]
fn anonymous_class_with_constructor() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new class("world") {
    public function __construct(private string $name) {}
    public function greet(): string {
        return "hello $this->name";
    }
};
echo $obj->greet(), "\n";
"#
        ),
        &["hello world"]
    );
}

#[test]
fn anonymous_class_implements_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Printable {
    public function toString(): string;
}
$obj = new class implements Printable {
    public function toString(): string {
        return "I am printable";
    }
};
echo $obj->toString(), "\n";
"#
        ),
        &["I am printable"]
    );
}

#[test]
fn anonymous_class_extends() {
    assert_eq!(
        run_prints(
            r#"<?php
class Logger {
    public function log(string $msg): void {
        echo "LOG: $msg", "\n";
    }
}
$obj = new class extends Logger {
    public function info(string $msg): void {
        $this->log("INFO: $msg");
    }
};
$obj->info("started");
"#
        ),
        &["LOG: INFO: started"]
    );
}

#[test]
fn anonymous_class_counter_state() {
    assert_eq!(
        run_prints(
            r#"<?php
function makeCounter(int $start = 0): object {
    return new class($start) {
        private int $value;
        public function __construct(int $start) { $this->value = $start; }
        public function inc(): void { $this->value++; }
        public function get(): int { return $this->value; }
    };
}
$c = makeCounter(10);
$c->inc();
$c->inc();
$c->inc();
echo $c->get(), "\n";
"#
        ),
        &["13"]
    );
}

// ── Trait conflict resolution ────────────────────────────────────
#[test]
fn trait_insteadof() {
    assert_eq!(
        run_prints(
            r#"<?php
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
echo $c->hello(), "\n";
"#
        ),
        &["A"]
    );
}

#[test]
fn trait_alias_as() {
    assert_eq!(
        run_prints(
            r#"<?php
trait Greeter {
    public function hello(): string { return "hello"; }
}
class MyClass {
    use Greeter {
        hello as greet;
    }
}
$obj = new MyClass();
echo $obj->greet(), "\n";
echo $obj->hello(), "\n";
"#
        ),
        &["hello", "hello"]
    );
}

#[test]
fn trait_insteadof_with_alias() {
    assert_eq!(
        run_prints(
            r#"<?php
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
echo $z->speak(), "\n";
echo $z->ySpeak(), "\n";
"#
        ),
        &["X speaks", "Y speaks"]
    );
}

// ── Abstract methods in traits ───────────────────────────────────
#[test]
fn trait_abstract_method() {
    assert_eq!(
        run_prints(
            r#"<?php
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
echo $e1->isValid(), "\n";
echo $e2->isValid(), "\n";
"#
        ),
        &["valid", "invalid"]
    );
}

// ── Trait with constants (PHP 8.2) ───────────────────────────────
#[test]
fn trait_with_constants() {
    assert_eq!(
        run_prints(
            r#"<?php
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
echo $app->getVersion(), "\n";
"#
        ),
        &["1.0"]
    );
}

// ── Interface constants and default methods ──────────────────────
#[test]
fn interface_with_constants() {
    assert_eq!(
        run_prints(
            r#"<?php
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
echo $u->getStatus(), "\n";
echo User::INACTIVE, "\n";
"#
        ),
        &["1", "0"]
    );
}

// ── Multiple interface implementation ────────────────────────────
#[test]
fn multiple_interfaces() {
    assert_eq!(
        run_prints(
            r#"<?php
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
echo $f->read(), "\n";
"#
        ),
        &["hello world"]
    );
}

// ── Abstract class with concrete methods ─────────────────────────
#[test]
fn abstract_template_method() {
    assert_eq!(
        run_prints(
            r#"<?php
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
echo $r->generate(), "\n";
"#
        ),
        &["Q1: 100, Q2: 200, Q3: 150"]
    );
}

// ── Covariant return types ───────────────────────────────────────
#[test]
fn covariant_return() {
    assert_eq!(
        run_prints(
            r#"<?php
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
echo $c->first(), "\n";
"#
        ),
        &["hello"]
    );
}

// ── Property hooks / accessors pattern ───────────────────────────
#[test]
fn getter_setter_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
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
echo $t->getCelsius(), "\n";
echo $t->getFahrenheit(), "\n";
"#
        ),
        &["100", "212"]
    );
}

// ── Enum implementing interface ──────────────────────────────────
#[test]
fn enum_implements_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
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
echo Color::Red->label(), "\n";
echo Color::Green->value, "\n";
"#
        ),
        &["RED", "green"]
    );
}

// ── Enum from / tryFrom ──────────────────────────────────────────
#[test]
fn enum_from_tryfrom() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Suit: string {
    case Hearts = "H";
    case Diamonds = "D";
    case Clubs = "C";
    case Spades = "S";
}
$s = Suit::from("H");
echo $s->name, "\n";
$t = Suit::tryFrom("X");
echo $t === null ? "null" : $t->name, "\n";
"#
        ),
        &["Hearts", "null"]
    );
}

// ── Readonly constructor promotion ───────────────────────────────
#[test]
fn readonly_promotion_combo() {
    assert_eq!(
        run_prints(
            r#"<?php
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
echo $a->distanceTo($b), "\n";
"#
        ),
        &["5"]
    );
}

// ── Intersection types ───────────────────────────────────────────
#[test]
fn intersection_type_param() {
    assert_eq!(
        run_prints(
            r#"<?php
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
    echo $obj->count(), "\n";
    echo $obj, "\n";
}
describe(new Items(["a", "b", "c"]));
"#
        ),
        &["3", "a,b,c"]
    );
}

// ── Named arguments ──────────────────────────────────────────────
#[test]
fn named_args_skip_defaults() {
    assert_eq!(
        run_prints(
            r#"<?php
function createTag(string $tag, string $content, string $class = "", string $id = ""): string {
    $attrs = "";
    if ($class) $attrs .= " class=\"$class\"";
    if ($id) $attrs .= " id=\"$id\"";
    return "<$tag$attrs>$content</$tag>";
}
echo createTag("div", "hello", id: "main"), "\n";
echo createTag(tag: "span", content: "world", class: "bold"), "\n";
"#
        ),
        &[
            "<div id=\"main\">hello</div>",
            "<span class=\"bold\">world</span>"
        ]
    );
}

// ── Stringable interface ─────────────────────────────────────────
#[test]
fn stringable_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
class Money implements Stringable {
    public function __construct(private int $cents) {}
    public function __toString(): string {
        return "$" . number_format($this->cents / 100, 2);
    }
}
function display(Stringable $item): void {
    echo $item, "\n";
}
display(new Money(1299));
display(new Money(50));
"#
        ),
        &["$12.99", "$0.50"]
    );
}

// ── Object cloning ───────────────────────────────────────────────
#[test]
fn object_clone_shallow() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    public string $env = "prod";
    public int $timeout = 30;
}
$a = new Config();
$b = clone $a;
$b->env = "dev";
echo $a->env, "\n";
echo $b->env, "\n";
echo $a->timeout, "\n";
"#
        ),
        &["prod", "dev", "30"]
    );
}

#[test]
fn clone_deep_copy_with_magic() {
    assert_eq!(
        run_prints(
            r#"<?php
class Address {
    public function __construct(public string $city) {}
}
class Person {
    public Address $address;
    public function __construct(public string $name, string $city) {
        $this->address = new Address($city);
    }
    public function __clone() {
        $this->address = clone $this->address;
    }
}
$alice = new Person("Alice", "Paris");
$bob = clone $alice;
$bob->name = "Bob";
$bob->address->city = "London";
echo $alice->name . ":" . $alice->address->city, "\n";
echo $bob->name . ":" . $bob->address->city, "\n";
"#
        ),
        &["Alice:Paris", "Bob:London"]
    );
}

#[test]
fn final_class_cannot_override_runtime_signal() {
    assert_eq!(
        run_prints(
            r#"<?php
class BaseFormatter {
    final public function format(string $value): string { return strtoupper($value); }
}
class AppFormatter extends BaseFormatter {
    public function title(string $value): string { return $this->format($value); }
}
echo (new AppFormatter())->title("ok");
"#,
        ),
        &["OK"]
    );
}

#[test]
fn object_final_property_behavior() {
    assert_eq!(
        run_prints(
            r#"<?php
class Packet {
    public readonly string $id;
    public function __construct(string $id) { $this->id = $id; }
}
$p = new Packet("p-1");
echo $p->id;
"#,
        ),
        &["p-1"]
    );
}

#[test]
fn trait_with_static_property_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
trait Counter {
    protected static int $count = 0;
    public function next(): int {
        return ++self::$count;
    }
}
class A {
    use Counter;
}
class B {
    use Counter;
}
$a = new A();
$b = new B();
echo $a->next();
echo $b->next();
"#,
        ),
        &["1", "2"]
    );
}

#[test]
fn magic_property_hooks_runtime_check() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box {
    private array $store = [];
    public function __set(string $name, mixed $value): void { $this->store[$name] = $value; }
    public function __get(string $name): mixed { return $this->store[$name] ?? null; }
    public function has(string $name): bool { return array_key_exists($name, $this->store); }
}
$b = new Box();
$b->x = 10;
echo $b->has('x') ? 'yes' : 'no';
echo $b->x;
"#,
        ),
        &["yes10"]
    );
}

#[test]
fn interface_runtime_dispatch_with_instanceof() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Processor {
    public function process(string $v): string;
}
class TrimProcessor implements Processor {
    public function process(string $v): string { return trim($v); }
}
function execute(Processor $processor, string $value): string {
    return $processor->process($value);
}
$p = new TrimProcessor();
echo ($p instanceof Processor) ? 'ok' : 'bad';
echo execute($p, "  x ");
"#,
        ),
        &["ok x "]
    );
}

// ── Object identity ──────────────────────────────────────────────
#[test]
fn object_identity_vs_equality() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box {
    public function __construct(public int $value) {}
}
$a = new Box(5);
$b = $a;
$c = new Box(5);
echo ($a === $b) ? "same" : "different", "\n";
echo ($a === $c) ? "same" : "different", "\n";
echo ($a == $c) ? "equal" : "not equal", "\n";
"#
        ),
        &["same", "different", "equal"]
    );
}

// ── Fluent builder pattern ───────────────────────────────────────
#[test]
fn fluent_builder_returns_this() {
    assert_eq!(
        run_prints(
            r#"<?php
class QueryBuilder {
    private string $table = "";
    private array $conditions = [];
    private ?int $limitVal = null;

    public function from(string $table): static {
        $this->table = $table;
        return $this;
    }
    public function where(string $cond): static {
        $this->conditions[] = $cond;
        return $this;
    }
    public function limit(int $n): static {
        $this->limitVal = $n;
        return $this;
    }
    public function build(): string {
        $sql = "SELECT * FROM {$this->table}";
        if ($this->conditions) {
            $sql .= " WHERE " . implode(" AND ", $this->conditions);
        }
        if ($this->limitVal !== null) {
            $sql .= " LIMIT {$this->limitVal}";
        }
        return $sql;
    }
}
$q = (new QueryBuilder())
    ->from("users")
    ->where("active=1")
    ->where("age>18")
    ->limit(10)
    ->build();
echo $q, "\n";
"#
        ),
        &["SELECT * FROM users WHERE active=1 AND age>18 LIMIT 10"]
    );
}

// ── Final class / method ─────────────────────────────────────────
#[test]
fn final_class_compile_ok() {
    compile_ok(
        r#"<?php
final class Singleton {
    private static ?self $instance = null;
    private function __construct(public readonly string $id) {}
    public static function getInstance(): self {
        if (self::$instance === null) {
            self::$instance = new self("main");
        }
        return self::$instance;
    }
}
$s = Singleton::getInstance();
echo $s->id, "\n";
"#,
    );
}

#[test]
fn final_method_in_hierarchy() {
    compile_ok(
        r#"<?php
class Base {
    final public function identity(): string {
        return static::class;
    }
    public function greeting(): string {
        return "Hello from " . $this->identity();
    }
}
class Child extends Base {
    public function greeting(): string {
        return parent::greeting() . " (child)";
    }
}
$c = new Child();
echo $c->greeting(), "\n";
"#,
    );
}

// ── Multiple levels of inheritance + parent chain ───────────────
#[test]
fn three_level_parent_construct_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
class A {
    protected string $log = "";
    public function __construct() {
        $this->log .= "A";
    }
}
class B extends A {
    public function __construct() {
        parent::__construct();
        $this->log .= "B";
    }
}
class C extends B {
    public function __construct() {
        parent::__construct();
        $this->log .= "C";
    }
    public function getLog(): string { return $this->log; }
}
$c = new C();
echo $c->getLog(), "\n";
"#
        ),
        &["ABC"]
    );
}

// ── Static property shared across instances ──────────────────────
#[test]
fn static_accumulator_across_instances() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    private static int $total = 0;
    private int $id;
    public function __construct() {
        self::$total++;
        $this->id = self::$total;
    }
    public function getId(): int { return $this->id; }
    public static function getTotal(): int { return self::$total; }
}
$a = new Counter();
$b = new Counter();
$c = new Counter();
echo $a->getId(), "\n";
echo $b->getId(), "\n";
echo $c->getId(), "\n";
echo Counter::getTotal(), "\n";
"#
        ),
        &["1", "2", "3", "3"]
    );
}

// ── Static property inheritance ──────────────────────────────────
#[test]
fn static_property_per_subclass() {
    assert_eq!(
        run_prints(
            r#"<?php
class Registry {
    protected static array $items = [];
    public static function add(string $item): void {
        static::$items[] = $item;
    }
    public static function all(): array {
        return static::$items;
    }
}
class FruitRegistry extends Registry {
    protected static array $items = [];
}
class VegRegistry extends Registry {
    protected static array $items = [];
}
FruitRegistry::add("apple");
FruitRegistry::add("banana");
VegRegistry::add("carrot");
echo implode(",", FruitRegistry::all()), "\n";
echo implode(",", VegRegistry::all()), "\n";
"#
        ),
        &["apple,banana", "carrot"]
    );
}

// ── Class constants with expressions ────────────────────────────
#[test]
fn class_constant_expression() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    const BASE = 10;
    const DOUBLE = self::BASE * 2;
    const LABEL = "max:" . self::DOUBLE;
}
echo Config::BASE, "\n";
echo Config::DOUBLE, "\n";
echo Config::LABEL, "\n";
"#
        ),
        &["10", "20", "max:20"]
    );
}

// ── Object cast to array ─────────────────────────────────────────
#[test]
fn object_cast_to_array() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point {
    public function __construct(
        public int $x,
        public int $y,
    ) {}
}
$p = new Point(3, 4);
$arr = (array) $p;
echo $arr["x"], "\n";
echo $arr["y"], "\n";
"#
        ),
        &["3", "4"]
    );
}

// ── get_class / is_a ────────────────────────────────────────────
#[test]
fn get_class_and_is_a() {
    assert_eq!(
        run_prints(
            r#"<?php
class Animal {}
class Dog extends Animal {}
$d = new Dog();
echo get_class($d), "\n";
echo is_a($d, "Dog") ? "yes" : "no", "\n";
echo is_a($d, "Animal") ? "yes" : "no", "\n";
echo is_a($d, "Cat") ? "yes" : "no", "\n";
"#
        ),
        &["Dog", "yes", "yes", "no"]
    );
}

// ── method_exists / property_exists ─────────────────────────────
#[test]
fn method_exists_and_property_exists() {
    assert_eq!(
        run_prints(
            r#"<?php
class Widget {
    public string $name = "btn";
    private int $id = 1;
    public function render(): string { return "<button>"; }
}
$w = new Widget();
echo method_exists($w, "render") ? "yes" : "no", "\n";
echo method_exists($w, "missing") ? "yes" : "no", "\n";
echo property_exists($w, "name") ? "yes" : "no", "\n";
echo property_exists($w, "id") ? "yes" : "no", "\n";
"#
        ),
        &["yes", "no", "yes", "yes"]
    );
}

// ── Union types ──────────────────────────────────────────────────
#[test]
fn union_type_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class Response {
    public int|string $code;
    public function __construct(int|string $code) {
        $this->code = $code;
    }
    public function isOk(): bool {
        return $this->code === 200 || $this->code === "ok";
    }
}
$r1 = new Response(200);
$r2 = new Response("ok");
$r3 = new Response(500);
echo $r1->isOk() ? "ok" : "fail", "\n";
echo $r2->isOk() ? "ok" : "fail", "\n";
echo $r3->isOk() ? "ok" : "fail", "\n";
"#
        ),
        &["ok", "ok", "fail"]
    );
}

// ── Nullable types ────────────────────────────────────────────────
#[test]
fn nullable_type_method_params() {
    assert_eq!(
        run_prints(
            r#"<?php
class User {
    public function __construct(
        public string $name,
        public ?string $email = null,
    ) {}
    public function contact(): string {
        return $this->email ?? "no email";
    }
}
$u1 = new User("Alice", "alice@example.com");
$u2 = new User("Bob");
echo $u1->contact(), "\n";
echo $u2->contact(), "\n";
"#
        ),
        &["alice@example.com", "no email"]
    );
}

// ── Immutable value object pattern ───────────────────────────────
#[test]
fn immutable_value_object_wither() {
    assert_eq!(
        run_prints(
            r#"<?php
class Money {
    public function __construct(
        public readonly int $amount,
        public readonly string $currency,
    ) {}
    public function withAmount(int $amount): self {
        return new self($amount, $this->currency);
    }
    public function withCurrency(string $currency): self {
        return new self($this->amount, $currency);
    }
    public function __toString(): string {
        return "{$this->amount} {$this->currency}";
    }
}
$m1 = new Money(100, "USD");
$m2 = $m1->withAmount(200);
$m3 = $m2->withCurrency("EUR");
echo $m1, "\n";
echo $m2, "\n";
echo $m3, "\n";
"#
        ),
        &["100 USD", "200 USD", "200 EUR"]
    );
}

// ── Class defined inside function (closure-like encapsulation) ───
#[test]
fn class_defined_inside_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function makeNode(int $value): object {
    class Node {
        public ?Node $next = null;
        public function __construct(public int $value) {}
    }
    return new Node($value);
}
$n = makeNode(42);
echo $n->value, "\n";
"#
        ),
        &["42"]
    );
}

// ── new static() vs new self() ────────────────────────────────────
#[test]
fn new_self_vs_new_static_in_clone_method() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    protected string $tag;
    public function __construct(string $tag) { $this->tag = $tag; }
    public function cloneSelf(): self   { return new self("base-copy"); }
    public function cloneStatic(): static { return new static($this->tag . "-copy"); }
    public function getTag(): string { return $this->tag; }
}
class Sub extends Base {}
$s = new Sub("sub");
$a = $s->cloneSelf();
$b = $s->cloneStatic();
echo get_class($a) . ":" . $a->getTag(), "\n";
echo get_class($b) . ":" . $b->getTag(), "\n";
"#
        ),
        &["Base:base-copy", "Sub:sub-copy"]
    );
}

// ── Method with splat args ────────────────────────────────────────
#[test]
fn variadic_method_collect_args() {
    assert_eq!(
        run_prints(
            r#"<?php
class Formatter {
    public function format(string $tpl, mixed ...$args): string {
        return vsprintf($tpl, $args);
    }
}
$f = new Formatter();
echo $f->format("%s is %d years old", "Alice", 30), "\n";
echo $f->format("%.2f + %.2f = %.2f", 1.1, 2.2, 3.3), "\n";
"#
        ),
        &["Alice is 30 years old", "1.10 + 2.20 = 3.30"]
    );
}

// ── Variadic method override in child ────────────────────────────
#[test]
fn variadic_override_in_child_class() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public function combine(string ...$parts): string {
        return implode("-", $parts);
    }
}
class Child extends Base {
    public function combine(string ...$parts): string {
        $upper = array_map("strtoupper", $parts);
        return parent::combine(...$upper);
    }
}
$c = new Child();
echo $c->combine("foo", "bar", "baz"), "\n";
"#
        ),
        &["FOO-BAR-BAZ"]
    );
}

// ── Object iteration via Iterator interface ───────────────────────
#[test]
fn object_implements_iterator() {
    assert_eq!(
        run_prints(
            r#"<?php
class NumberRange implements Iterator {
    private int $current;
    public function __construct(
        private int $start,
        private int $end,
    ) {
        $this->current = $start;
    }
    public function current(): int  { return $this->current; }
    public function key(): int      { return $this->current - $this->start; }
    public function next(): void    { $this->current++; }
    public function rewind(): void  { $this->current = $this->start; }
    public function valid(): bool   { return $this->current <= $this->end; }
}
$range = new NumberRange(1, 5);
$vals = [];
foreach ($range as $k => $v) {
    $vals[] = "$k:$v";
}
echo implode(",", $vals), "\n";
"#
        ),
        &["0:1,1:2,2:3,3:4,4:5"]
    );
}

// ── Object in array_map ──────────────────────────────────────────
#[test]
fn objects_in_array_map() {
    assert_eq!(
        run_prints(
            r#"<?php
class Item {
    public function __construct(public string $name, public float $price) {}
    public function discounted(float $pct): self {
        return new self($this->name, $this->price * (1 - $pct));
    }
}
$items = [
    new Item("Widget", 10.0),
    new Item("Gadget", 20.0),
    new Item("Doohickey", 5.0),
];
$discounted = array_map(fn($i) => $i->discounted(0.1), $items);
$totals = array_map(fn($i) => $i->price, $discounted);
echo number_format(array_sum($totals), 2), "\n";
"#
        ),
        &["31.50"]
    );
}

// ── Named constructor / static factory returning subtype ─────────
#[test]
fn named_constructor_static_factory() {
    assert_eq!(
        run_prints(
            r#"<?php
class Color {
    private function __construct(
        private int $r,
        private int $g,
        private int $b,
    ) {}
    public static function fromHex(string $hex): self {
        $hex = ltrim($hex, '#');
        return new self(
            hexdec(substr($hex, 0, 2)),
            hexdec(substr($hex, 2, 2)),
            hexdec(substr($hex, 4, 2)),
        );
    }
    public static function fromRgb(int $r, int $g, int $b): self {
        return new self($r, $g, $b);
    }
    public function __toString(): string {
        return "rgb({$this->r},{$this->g},{$this->b})";
    }
}
$c1 = Color::fromHex('#ff8000');
$c2 = Color::fromRgb(0, 128, 255);
echo $c1, "\n";
echo $c2, "\n";
"#
        ),
        &["rgb(255,128,0)", "rgb(0,128,255)"]
    );
}

// ── Interface extending multiple interfaces ───────────────────────
#[test]
fn interface_extends_multiple() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Serializable2 {
    public function serialize(): string;
}
interface Deserializable {
    public static function deserialize(string $data): static;
}
interface Codec extends Serializable2, Deserializable {}
class JsonRecord implements Codec {
    public array $data;
    public function __construct(array $data) { $this->data = $data; }
    public function serialize(): string { return json_encode($this->data); }
    public static function deserialize(string $data): static {
        return new static(json_decode($data, true));
    }
}
$r = new JsonRecord(["key" => "value"]);
$s = $r->serialize();
$r2 = JsonRecord::deserialize($s);
echo $s, "\n";
echo $r2->data["key"], "\n";
"#
        ),
        &["{\"key\":\"value\"}", "value"]
    );
}

// ── Recursive method calling parent and child ────────────────────
#[test]
fn recursive_method_with_inheritance() {
    assert_eq!(
        run_prints(
            r#"<?php
class TreeNode {
    public ?TreeNode $left = null;
    public ?TreeNode $right = null;
    public function __construct(public int $value) {}
    public function insert(int $v): void {
        if ($v < $this->value) {
            if ($this->left === null) $this->left = new static($v);
            else $this->left->insert($v);
        } else {
            if ($this->right === null) $this->right = new static($v);
            else $this->right->insert($v);
        }
    }
    public function inorder(): array {
        $result = [];
        if ($this->left !== null) $result = array_merge($result, $this->left->inorder());
        $result[] = $this->value;
        if ($this->right !== null) $result = array_merge($result, $this->right->inorder());
        return $result;
    }
}
$tree = new TreeNode(5);
foreach ([3, 7, 1, 4, 6, 8] as $v) {
    $tree->insert($v);
}
echo implode(",", $tree->inorder()), "\n";
"#
        ),
        &["1,3,4,5,6,7,8"]
    );
}

// ── Cloning with array property deep copy ────────────────────────
#[test]
fn clone_deep_copy_array_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class ShoppingCart {
    private array $items = [];
    public function add(string $item): void {
        $this->items[] = $item;
    }
    public function __clone() {
        // array is value-type so no manual deep copy needed,
        // but let's verify mutation isolation
        $this->items = $this->items; // explicit reassign
    }
    public function count(): int { return count($this->items); }
    public function items(): array { return $this->items; }
}
$cart1 = new ShoppingCart();
$cart1->add("apple");
$cart1->add("banana");
$cart2 = clone $cart1;
$cart2->add("cherry");
echo $cart1->count(), "\n";
echo $cart2->count(), "\n";
echo implode(",", $cart1->items()), "\n";
"#
        ),
        &["2", "3", "apple,banana"]
    );
}

// ── Interface without constructor ─────────────────────────────────
#[test]
fn interface_has_no_constructor() {
    compile_ok(
        r#"<?php
interface Shape {
    public function area(): float;
    public function perimeter(): float;
}
class Rect implements Shape {
    public function __construct(private float $w, private float $h) {}
    public function area(): float { return $this->w * $this->h; }
    public function perimeter(): float { return 2 * ($this->w + $this->h); }
}
$r = new Rect(3.0, 4.0);
echo $r->area(), "\n";
echo $r->perimeter(), "\n";
"#,
    );
}

// ── Mixed type in method signature ───────────────────────────────
#[test]
fn mixed_type_in_method_signature() {
    assert_eq!(
        run_prints(
            r#"<?php
class Converter {
    public function toInt(mixed $value): int {
        return (int) $value;
    }
    public function toBool(mixed $value): bool {
        return (bool) $value;
    }
    public function toStr(mixed $value): string {
        return (string) $value;
    }
}
$c = new Converter();
echo $c->toInt("42"), "\n";
echo $c->toBool(0) ? "true" : "false", "\n";
echo $c->toStr(3.14), "\n";
"#
        ),
        &["42", "false", "3.14"]
    );
}

// ── Abstract class with multiple abstract methods ─────────────────
#[test]
fn abstract_class_multiple_abstract_methods() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Serializer {
    abstract protected function encode(array $data): string;
    abstract protected function decode(string $raw): array;
    public function roundtrip(array $data): array {
        return $this->decode($this->encode($data));
    }
}
class JsonSerializer extends Serializer {
    protected function encode(array $data): string { return json_encode($data); }
    protected function decode(string $raw): array { return json_decode($raw, true); }
}
$s = new JsonSerializer();
$result = $s->roundtrip(["x" => 1, "y" => 2]);
echo $result["x"], "\n";
echo $result["y"], "\n";
"#
        ),
        &["1", "2"]
    );
}

// ── Interface dispatch plus instanceof in one runtime path ───────
#[test]
fn interface_dispatch_with_runtime_check() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Transform {
    public function apply(string $value): string;
}
class JsonTransform implements Transform {
    public function apply(string $value): string {
        return json_encode(["value" => $value]);
    }
}
$transform = new JsonTransform();
echo ($transform instanceof Transform) ? "yes" : "no";
echo "|";
echo $transform->apply("x"), "\n";
"#
        ),
        &["yes|\"value\":\"x\""]
    );
}

// ── Named constructor and protected constructor mixin ───────────
#[test]
fn protected_constructor_with_static_factory() {
    assert_eq!(
        run_prints(
            r#"<?php
class Node {
    protected function __construct(public int $value) {}
    public static function from(int $value): self {
        return new self($value);
    }
    public function double(): int { return $this->value * 2; }
}
echo Node::from(7)->double(), "\n";
"#
        ),
        &["14"]
    );
}

#[test]
fn trait_private_property_visibility_with_public_method() {
    assert_eq!(
        run_prints(
            r#"<?php
trait Store {
    private string $value = "x";
    public function value(): string { return $this->value; }
}
class Holder {
    use Store;
}
echo (new Holder())->value(), "\n";
"#
        ),
        &["x"]
    );
}

#[test]
fn method_exists_runtime_on_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box {
    public function open(): void {}
}
$box = new Box();
echo method_exists($box, "open") ? "yes" : "no";
echo method_exists($box, "close") ? "yes" : "no";
"#
        ),
        &["yesno"]
    );
}

#[test]
fn dynamic_method_call_on_object() {
    assert_eq!(
        run_prints(
            r#"<?php
class Worker {
    public function run(string $task): string { return "run:$task"; }
    public function stop(): string { return "stop"; }
}
$w = new Worker();
$method = "run";
echo $w->$method("build");
echo "|";
echo $w->{"stop"}();
"#
        ),
        &["run:build|stop"]
    );
}

#[test]
fn static_method_dynamic_name_call() {
    assert_eq!(
        run_prints(
            r#"<?php
class Codec {
    public static function encode(string $v): string { return strtoupper($v); }
    public static function decode(string $v): string { return strtolower($v); }
}
$step = "encode";
$class = Codec::class;
echo $class::$step("ab");
echo "|";
$step = "decode";
echo $class::$step("XY");
"#
        ),
        &["AB|xy"]
    );
}

#[test]
fn chaining_with_arrow_and_returned_object() {
    assert_eq!(
        run_prints(
            r#"<?php
class Chain {
    public int $value = 1;
    public function inc(int $n): self {
        $this->value += $n;
        return $this;
    }
    public function scale(int $n): self {
        $this->value *= $n;
        return $this;
    }
}
$c = new Chain();
echo $c->inc(2)->scale(3)->value;
"#
        ),
        &["9"]
    );
}

#[test]
fn magic_call_allows_deferred_methods() {
    assert_eq!(
        run_prints(
            r#"<?php
class Broker {
    public function __call(string $name, array $args): string {
        return $name . ":" . implode(",", $args);
    }
}
$b = new Broker();
echo $b->enqueue("a", 1);
echo "|";
echo $b->flush();
"#
        ),
        &["enqueue:a,1|flush:"]
    );
}

#[test]
fn readonly_property_assignment_raises_error() {
    assert_eq!(
        run_prints(
            r#"<?php
class Snapshot {
    public readonly string $name;
    public function __construct(string $name) { $this->name = $name; }
}
$s = new Snapshot("init");
try {
    $s->name = "bad";
    echo "changed";
} catch (Error $e) {
    echo "readonly";
}
"#
        ),
        &["readonly"]
    );
}
