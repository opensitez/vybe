use super::helpers::compile_ok;

// ── Covariant return type — child returns narrower type ───────────
#[test]
fn covariant_return_type() {
    compile_ok(
        r#"<?php
class Animal {}
class Dog extends Animal {}
class AnimalFactory {
    public function create(): Animal { return new Animal(); }
}
class DogFactory extends AnimalFactory {
    public function create(): Dog { return new Dog(); }
}
$f = new DogFactory();
echo $f->create() instanceof Dog ? 'dog' : 'not-dog';
"#,
    );
}

// ── Object comparison == vs === ───────────────────────────────────
#[test]
fn object_comparison_loose_vs_strict() {
    compile_ok(
        r#"<?php
class Point {
    public function __construct(public int $x, public int $y) {}
}
$a = new Point(1, 2);
$b = new Point(1, 2);
$c = $a;
echo ($a == $b)  ? 'loose-eq'  : 'loose-neq';
echo ($a === $b) ? 'strict-eq' : 'strict-neq';
echo ($a === $c) ? 'strict-eq' : 'strict-neq';
"#,
    );
}

// ── Cloning with __clone for deep copy ───────────────────────────
#[test]
fn clone_with_deep_copy() {
    compile_ok(
        r#"<?php
class Address {
    public function __construct(public string $city) {}
}
class Person {
    public function __construct(public string $name, public Address $address) {}
    public function __clone() {
        $this->address = clone $this->address;
    }
}
$original = new Person('Alice', new Address('Paris'));
$copy = clone $original;
$copy->address->city = 'London';
echo $original->address->city;
echo $copy->address->city;
"#,
    );
}

// ── Static property as shared state across instances ──────────────
#[test]
fn static_property_shared_state() {
    compile_ok(
        r#"<?php
class Counter {
    private static int $count = 0;
    public function __construct() { self::$count++; }
    public static function total(): int { return self::$count; }
}
new Counter();
new Counter();
new Counter();
echo Counter::total();
"#,
    );
}

// ── Constant in interface implemented by class ────────────────────
#[test]
fn interface_constant_implementation() {
    compile_ok(
        r#"<?php
interface HasVersion {
    const API_VERSION = '2.0';
}
class Client implements HasVersion {
    public function version(): string { return self::API_VERSION; }
}
$c = new Client();
echo $c->version();
echo Client::API_VERSION;
"#,
    );
}

// ── Multiple traits in one class with method from each ───────────
#[test]
fn multiple_traits_methods_from_each() {
    compile_ok(
        r#"<?php
trait Serializable {
    public function serialize(): string { return json_encode($this->toArray()); }
}
trait Loggable {
    public function log(): string { return get_class($this) . ':logged'; }
}
class Order {
    use Serializable, Loggable;
    public function __construct(private int $id, private float $total) {}
    public function toArray(): array { return ['id' => $this->id, 'total' => $this->total]; }
}
$o = new Order(42, 99.99);
echo $o->log();
echo is_string($o->serialize()) ? 'serialized' : 'fail';
"#,
    );
}

// ── Trait conflict resolution with insteadof ──────────────────────
#[test]
fn trait_insteadof_conflict_resolution() {
    compile_ok(
        r#"<?php
trait A {
    public function hello(): string { return 'A::hello'; }
}
trait B {
    public function hello(): string { return 'B::hello'; }
}
class MyClass {
    use A, B {
        A::hello insteadof B;
        B::hello as helloFromB;
    }
}
$obj = new MyClass();
echo $obj->hello();
echo $obj->helloFromB();
"#,
    );
}

// ── Abstract class with abstract and concrete methods ────────────
#[test]
fn abstract_class_mixed_methods() {
    compile_ok(
        r#"<?php
abstract class Template {
    abstract protected function step1(): string;
    abstract protected function step2(): string;
    public function run(): string {
        return $this->step1() . '|' . $this->step2();
    }
    public function describe(): string { return 'Template'; }
}
class ConcreteTemplate extends Template {
    protected function step1(): string { return 'init'; }
    protected function step2(): string { return 'execute'; }
}
$t = new ConcreteTemplate();
echo $t->run();
echo $t->describe();
"#,
    );
}

// ── Final method cannot be overridden (structural) ────────────────
#[test]
fn final_method_structural() {
    compile_ok(
        r#"<?php
class Base {
    final public function identity(): string { return static::class; }
    public function greet(): string { return 'hello from ' . $this->identity(); }
}
class Child extends Base {
    public function greet(): string { return 'child greet: ' . $this->identity(); }
}
$b = new Base();
$c = new Child();
echo $b->greet();
echo $c->greet();
"#,
    );
}

// ── Named constructor (static factory method) ─────────────────────
#[test]
fn named_constructor_factory() {
    compile_ok(
        r#"<?php
class Color {
    private function __construct(
        private int $r,
        private int $g,
        private int $b
    ) {}
    public static function fromRGB(int $r, int $g, int $b): self {
        return new self($r, $g, $b);
    }
    public static function fromHex(string $hex): self {
        $hex = ltrim($hex, '#');
        return new self(
            hexdec(substr($hex, 0, 2)),
            hexdec(substr($hex, 2, 2)),
            hexdec(substr($hex, 4, 2))
        );
    }
    public function toCSS(): string { return "rgb({$this->r},{$this->g},{$this->b})"; }
}
$red  = Color::fromRGB(255, 0, 0);
echo $red->toCSS();
"#,
    );
}

// ── Constructor with multiple optional parameters ─────────────────
#[test]
fn constructor_optional_parameters() {
    compile_ok(
        r#"<?php
class HttpRequest {
    public function __construct(
        public readonly string $method  = 'GET',
        public readonly string $path    = '/',
        public readonly array  $headers = [],
        public readonly string $body    = ''
    ) {}
}
$get    = new HttpRequest();
$post   = new HttpRequest('POST', '/submit', ['Content-Type' => 'application/json'], '{}');
echo $get->method . ' ' . $get->path;
echo $post->method . ' ' . $post->path;
"#,
    );
}

// ── Method chaining returning $this ──────────────────────────────
#[test]
fn method_chaining_returns_this() {
    compile_ok(
        r#"<?php
class QueryBuilder {
    private string $table  = '';
    private array  $wheres = [];
    private ?int   $limit  = null;
    public function from(string $t): static { $this->table = $t; return $this; }
    public function where(string $cond): static { $this->wheres[] = $cond; return $this; }
    public function limit(int $n): static { $this->limit = $n; return $this; }
    public function toSql(): string {
        $sql = 'SELECT * FROM ' . $this->table;
        if ($this->wheres) $sql .= ' WHERE ' . implode(' AND ', $this->wheres);
        if ($this->limit !== null) $sql .= ' LIMIT ' . $this->limit;
        return $sql;
    }
}
$q = (new QueryBuilder())->from('users')->where('active=1')->where('age>18')->limit(10);
echo $q->toSql();
"#,
    );
}

// ── Private constructor with public factory (singleton shape) ─────
#[test]
fn private_constructor_factory_pattern() {
    compile_ok(
        r#"<?php
class Database {
    private static ?self $instance = null;
    private function __construct(private string $dsn) {}
    public static function connect(string $dsn): self {
        if (self::$instance === null) {
            self::$instance = new self($dsn);
        }
        return self::$instance;
    }
    public function getDsn(): string { return $this->dsn; }
}
$db1 = Database::connect('mysql://localhost/app');
$db2 = Database::connect('ignored-because-already-connected');
echo $db1->getDsn();
echo ($db1 === $db2) ? 'same' : 'different';
"#,
    );
}

// ── Object used as SplObjectStorage key (structural compile_ok) ───
#[test]
fn object_as_storage_key_structural() {
    compile_ok(
        r#"<?php
class Token {
    public function __construct(public string $value) {}
}
$storage = new SplObjectStorage();
$t1 = new Token('abc');
$t2 = new Token('xyz');
$storage->attach($t1, 'meta-for-abc');
$storage->attach($t2, 'meta-for-xyz');
echo $storage->count();
echo $storage[$t1];
"#,
    );
}

// ── Nullsafe operator on chained object access ────────────────────
#[test]
fn nullsafe_operator_chained() {
    compile_ok(
        r#"<?php
class Street {
    public function __construct(public string $name) {}
}
class Address {
    public function __construct(public ?Street $street = null) {}
    public function getStreet(): ?Street { return $this->street; }
}
class User {
    public function __construct(public ?Address $address = null) {}
    public function getAddress(): ?Address { return $this->address; }
}
$userWithAddress    = new User(new Address(new Street('Main St')));
$userWithoutAddress = new User(null);
echo $userWithAddress?->getAddress()?->getStreet()?->name ?? 'none';
echo $userWithoutAddress?->getAddress()?->getStreet()?->name ?? 'none';
"#,
    );
}

// ── Calling parent constructor from child ─────────────────────────
#[test]
fn parent_constructor_call() {
    compile_ok(
        r#"<?php
class Vehicle {
    public function __construct(
        protected string $make,
        protected int    $year
    ) {}
    public function info(): string { return "{$this->year} {$this->make}"; }
}
class Car extends Vehicle {
    public function __construct(
        string $make,
        int    $year,
        private int $doors
    ) {
        parent::__construct($make, $year);
    }
    public function describe(): string { return $this->info() . " ({$this->doors} doors)"; }
}
$car = new Car('Toyota', 2023, 4);
echo $car->describe();
"#,
    );
}

// ── Interface type hint accepting any implementation ──────────────
#[test]
fn interface_typehint_polymorphism() {
    compile_ok(
        r#"<?php
interface Formatter {
    public function format(mixed $value): string;
}
class NumberFormatter implements Formatter {
    public function format(mixed $value): string { return number_format((float)$value, 2); }
}
class UpperFormatter implements Formatter {
    public function format(mixed $value): string { return strtoupper((string)$value); }
}
function render(Formatter $fmt, mixed $val): void {
    echo $fmt->format($val);
}
render(new NumberFormatter(), 1234.5);
render(new UpperFormatter(), 'hello');
"#,
    );
}

// ── Class implementing interface — instanceof check ───────────────
#[test]
fn instanceof_interface_check() {
    compile_ok(
        r#"<?php
interface Runnable {
    public function run(): void;
}
interface Stoppable {
    public function stop(): void;
}
class Engine implements Runnable, Stoppable {
    private bool $running = false;
    public function run(): void  { $this->running = true; }
    public function stop(): void { $this->running = false; }
    public function isRunning(): bool { return $this->running; }
}
$e = new Engine();
echo $e instanceof Runnable   ? 'runnable'   : 'not';
echo $e instanceof Stoppable  ? 'stoppable'  : 'not';
"#,
    );
}

// ── Magic __toString returning useful representation ──────────────
#[test]
fn magic_tostring_representation() {
    compile_ok(
        r#"<?php
class Vector2D {
    public function __construct(private float $x, private float $y) {}
    public function __toString(): string {
        return "({$this->x}, {$this->y})";
    }
    public function add(Vector2D $other): self {
        return new self($this->x + $other->x, $this->y + $other->y);
    }
}
$v1 = new Vector2D(1.0, 2.0);
$v2 = new Vector2D(3.0, 4.0);
echo $v1;
echo $v1->add($v2);
"#,
    );
}

// ── Typed property declaration with default value ─────────────────
#[test]
fn typed_property_with_default() {
    compile_ok(
        r#"<?php
class Config {
    public string   $env      = 'production';
    public int      $maxRetry = 3;
    public bool     $debug    = false;
    public float    $timeout  = 30.0;
    public array    $tags     = [];
    public ?string  $secret   = null;
}
$c = new Config();
echo $c->env;
echo $c->maxRetry;
echo $c->debug ? 'true' : 'false';
echo $c->timeout;
echo $c->secret ?? 'null';
"#,
    );
}
