use super::helpers::{compile_ok, run_prints};

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

#[test]
fn covariant_return_type_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Animal {}
class Dog extends Animal {}
class AnimalFactory { public function create(): Animal { return new Animal(); } }
class DogFactory extends AnimalFactory {
    public function create(): Dog { return new Dog(); }
}
$f = new DogFactory();
echo $f->create() instanceof Dog ? 'yes' : 'no';
"#,
        ),
        &["yes"]
    );
}

#[test]
fn object_comparison_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point {
    public function __construct(public int $x, public int $y) {}
}
$a = new Point(1, 2);
$b = new Point(1, 2);
$c = $a;
echo ($a == $b) ? 'e' : 'n';
echo ($a === $c) ? 's' : 'd';
"#,
        ),
        &["es"]
    );
}

#[test]
fn deep_clone_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Address {
    public function __construct(public string $city) {}
}
class Person {
    public function __construct(public string $name, public Address $address) {}
    public function __clone(): void {
        $this->address = clone $this->address;
    }
}
$original = new Person('Alice', new Address('Paris'));
$copy = clone $original;
$copy->address->city = 'London';
echo $original->address->city . '|' . $copy->address->city;
"#,
        ),
        &["Paris|London"]
    );
}

#[test]
fn static_property_shared_state_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    private static int $count = 0;
    public function __construct() { self::$count++; }
    public static function total(): int { return self::$count; }
}
new Counter();
new Counter();
echo (string) Counter::total();
"#,
        ),
        &["2"]
    );
}

#[test]
fn private_constructor_factory_runtime() {
    assert_eq!(
        run_prints(
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
$db2 = Database::connect('other://ignore');
echo $db1->getDsn() . '|' . (($db1 === $db2) ? 'same' : 'diff');
"#,
        ),
        &["mysql://localhost/app|same"]
    );
}

#[test]
fn trait_conflict_alias_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
trait A { public function label(): string { return 'A'; } }
trait B { public function label(): string { return 'B'; } }
class C {
    use A, B { B::label insteadof A; A::label as labelFromA; }
}
echo (new C())->label() . '|' . (new C())->labelFromA();
"#,
        ),
        &["B|A"]
    );
}

#[test]
fn nullsafe_chain_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Street { public function __construct(public string $name) {} }
class Address { public function __construct(public ?Street $street = null) {} }
class User {
    public function __construct(public ?Address $address = null) {}
}
echo (new User(new Address(new Street('Main')))?->address?->street?->name ?? 'none'), '|', (new User())->address?->street?->name ?? 'none';
"#,
        ),
        &["Main|none"]
    );
}

#[test]
fn dynamic_class_instantiation_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Service {
    public function __construct(public string $name) {}
}
$class = Service::class;
$svc = new $class('Search');
echo $svc->name;
"#,
        ),
        &["Search"]
    );
}

#[test]
fn dynamic_method_invocation_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Worker {
    public function work(string $task): string { return "work:$task"; }
}
$obj = new Worker();
$m = 'work';
echo $obj->$m('backup');
"#,
        ),
        &["work:backup"]
    );
}

#[test]
fn trait_aliasing_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
trait Logger {
    public function log(string $msg): string { return "log:$msg"; }
}
trait Auditor {
    public function log(string $msg): string { return "audit:$msg"; }
}
class ServiceLayer {
    use Logger, Auditor { Logger::log insteadof Auditor; Auditor::log as auditLog; }
}
$svc = new ServiceLayer();
echo $svc->log('ok');
echo $svc->auditLog('ok');
"#,
        ),
        &["log:okaudit:ok"]
    );
}

#[test]
fn readonly_visibility_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config {
    public function __construct(public readonly string $name, public readonly int $retries) {}
}
$c = new Config('core', 5);
echo $c->name;
echo $c->retries;
"#,
        ),
        &["core5"]
    );
}

#[test]
fn magic_invoke_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class CallableCounter {
    public function __construct(private int $n = 0) {}
    public function __invoke(int $step = 1): int {
        $this->n += $step;
        return $this->n;
    }
}
$c = new CallableCounter(2);
echo $c();
echo $c(3);
"#,
        ),
        &["2","5"]
    );
}

#[test]
fn interface_dispatch_with_multiple_implementers_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Formatter {
    public function format(string $value): string;
}

class Upper implements Formatter {
    public function format(string $value): string { return strtoupper($value); }
}

class Lower implements Formatter {
    public function format(string $value): string { return strtolower($value); }
}

function apply_formatter(Formatter $formatter, string $value): string {
    return $formatter->format($value);
}

echo apply_formatter(new Upper(), 'ab');
echo apply_formatter(new Lower(), 'AB');
"#,
        ),
        &["AB", "ab"]
    );
}

#[test]
fn late_static_factory_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class BaseProduct {
    public function __construct(public string $name) {}
    public static function make(string $name): static {
        return new static($name);
    }
}

class Widget extends BaseProduct {}

$w = Widget::make('widget');
echo get_class($w) . '|' . $w->name;
"#,
        ),
        &["Widget|widget"]
    );
}

#[test]
fn dynamic_class_property_exists_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Profile {
    public function __construct(public string $name = 'anon') {}
}

$class = 'Profile';
$property = 'name';
echo property_exists(new $class(), $property) ? 'yes' : 'no';
"#,
        ),
        &["yes"]
    );
}

#[test]
fn __call_forwarding_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Invoker {
    public function __call(string $name, array $args): string {
        return $name . ':' . implode(',', $args);
    }
}
$i = new Invoker();
echo $i->run('build', 1, 2);
"#,
        ),
        &["run:build,1,2"]
    );
}

#[test]
fn __call_static_and_property_visibility_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Service {
    private static int $counter = 0;
    public static function bump(): int {
        return ++self::$counter;
    }
}

echo Service::bump();
echo Service::bump();
"#,
        ),
        &["12"]
    );
}

#[test]
fn class_alias_and_instanceof_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Adapter {}
class_alias(Adapter::class, 'ServiceAdapter');
echo class_exists('ServiceAdapter') ? 'exists' : 'missing';
echo '|';
echo (new ServiceAdapter()) instanceof Adapter ? 'yes' : 'no';
"#,
        ),
        &["exists|yes"]
    );
}

#[test]
fn magic_setter_getter_with_isset_unset_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Bag {
    private array $values = [];
    public function __set(string $name, mixed $value): void {
        $this->values[$name] = $value;
    }
    public function __get(string $name): mixed {
        return $this->values[$name] ?? null;
    }
    public function __isset(string $name): bool {
        return array_key_exists($name, $this->values);
    }
    public function __unset(string $name): void {
        unset($this->values[$name]);
    }
}
$bag = new Bag();
$bag->token = 'abc';
echo $bag->token;
echo '|';
echo isset($bag->token) ? 'set' : 'not';
unset($bag->token);
echo '|';
echo isset($bag->token) ? 'set' : 'not';
"#,
        ),
        &["abc|set|not"]
    );
}

#[test]
fn object_set_state_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Payload {
    public string $value = '';
    public static function __set_state(array $state): self {
        $obj = new self();
        $obj->value = strtoupper($state['value']);
        return $obj;
    }
}
$obj = Payload::__set_state(['value' => 'ok']);
echo $obj->value;
"#,
        ),
        &["OK"]
    );
}

#[test]
fn serialize_cycle_with_unserialize_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box {
    public function __construct(public int $n, public string $label) {}
    public function __serialize(): array {
        return ['n' => $this->n, 'label' => $this->label];
    }
    public function __unserialize(array $data): void {
        $this->n = $data['n'];
        $this->label = $data['label'] . '!'; 
    }
}
$box = new Box(7, 'ok');
$text = serialize($box);
$copy = unserialize($text);
echo $copy->n . '|' . $copy->label;
"#,
        ),
        &["7|ok!"]
    );
}

#[test]
fn named_arguments_to_constructor_and_method_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Metrics {
    public function __construct(public int $a = 0, public int $b = 0) {}
    public function span(int $start = 0, int $end = 0): int {
        return $end - $start;
    }
}
$m = new Metrics(b: 30, a: 10);
echo $m->span(end: 15, start: 5) . '|' . $m->a . ':' . $m->b;
"#,
        ),
        &["10|10:30"]
    );
}

#[test]
fn late_static_binding_clone_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class BaseProduct {
    public static string $kind = 'base';
    public static function make(string $name): self {
        return new static($name);
    }
    public function __construct(public string $name) {}
    public function type(): string {
        return static::$kind;
    }
}
class PremiumProduct extends BaseProduct {
    public static string $kind = 'premium';
}
$product = PremiumProduct::make('x');
echo $product->type();
echo '|';
echo $product instanceof PremiumProduct ? 'premium' : 'base';
"#,
        ),
        &["premium|premium"]
    );
}

#[test]
fn readonly_property_reassignment_error_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Token {
    public function __construct(public readonly int $id) {}
}
$t = new Token(1);
$result = null;
try {
    $t->id = 2;
    $result = 'changed';
} catch (Error $e) {
    $result = 'readonly';
}
echo $t->id . '|' . $result;
"#,
        ),
        &["1|readonly"]
    );
}

#[test]
fn property_visibility_across_hierarchy_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public string $public = 'pub';
    protected string $protected = 'prot';
    private string $private = 'priv';
    public function visible(): string {
        return $this->public . '|' . $this->protected;
    }
}
class Child extends Base {
    public function secret(): string {
        return $this->protected;
    }
}
$obj = new Child();
echo $obj->visible();
echo '|' . $obj->secret();
echo '|' . $obj->public;
"#,
        ),
        &["pub|prot|pub|prot"]
    );
}

#[test]
fn class_const_visibility_and_inheritance_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class BaseValue {
    public const SCOPE = 'public';
    protected const INTERNAL = 'internal';
    private const SECRET = 'secret';
    public function marker(): string {
        return self::SCOPE . '|' . static::SCOPE;
    }
}
class ChildValue extends BaseValue {}
echo ChildValue::SCOPE;
echo '|' . (new ChildValue())->marker();
"#,
        ),
        &["public|public|public"]
    );
}

#[test]
fn dynamic_property_with_overloaded_setter_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Bag {
    private array $values = [];
    public function __set(string $name, mixed $value): void { $this->values[$name] = $value; }
    public function __get(string $name): mixed { return $this->values[$name] ?? null; }
}
$b = new Bag();
$b->one = 1;
$b->two = '2';
echo $b->one . '|' . $b->two;
"#,
        ),
        &["1|2"]
    );
}

#[test]
fn __invoke_on_object_with_union_input_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Processor {
    public function __invoke(string|int $v): string {
        return (string)$v;
    }
}
$p = new Processor();
echo $p('id');
echo '|';
echo $p(12);
"#,
        ),
        &["id|12"]
    );
}

#[test]
fn static_method_polymorphism_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class FormatterBase {
    public static function supports(): string { return 'base'; }
}
class JsonFormatter extends FormatterBase {
    public static function supports(): string { return 'json'; }
}
class CsvFormatter extends FormatterBase {
    public static function supports(): string { return 'csv'; }
}
echo JsonFormatter::supports();
echo '|' . CsvFormatter::supports();
echo '|' . FormatterBase::supports();
"#,
        ),
        &["json|csv|base"]
    );
}

#[test]
fn dynamic_type_check_with_instanceof_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface I {}
class A implements I {}
class B {}
$a = new A();
$b = new B();
echo ($a instanceof I ? 'ia' : 'noa');
echo '|' . ($b instanceof I ? 'ib' : 'nob');
"#,
        ),
        &["ia|nob"]
    );
}

#[test]
fn constructor_property_promotion_with_visibility_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class User {
    public function __construct(
        public string $name,
        protected int $age,
        private bool $active
    ) {}
    public function summary(): string {
        return $this->name . ':' . $this->age . ':' . ($this->active ? 'on' : 'off');
    }
}
$u = new User('alice', 31, true);
echo $u->name;
echo '|' . $u->summary();
"#,
        ),
        &["alice|alice:31:on"]
    );
}
