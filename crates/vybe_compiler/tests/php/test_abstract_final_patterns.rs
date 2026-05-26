use super::helpers::{compile_ok, run_prints};

// ── Abstract class basics ────────────────────────────────────

#[test] fn abstract_class_cannot_be_instantiated() {
    assert_eq!(run_prints(r#"<?php
abstract class Shape {}
try {
    new Shape();
} catch (Error $e) {
    echo "cannot instantiate";
}
"#), vec!["cannot instantiate"]);
}

#[test] fn abstract_method_must_be_implemented_in_child() {
    assert_eq!(run_prints(r#"<?php
abstract class Formatter {
    abstract public function format(string $s): string;
    public function process(string $s): string { return $this->format($s); }
}
class UpperFormatter extends Formatter {
    public function format(string $s): string { return strtoupper($s); }
}
echo (new UpperFormatter())->process("hello");
"#), vec!["HELLO"]);
}

#[test] fn abstract_class_with_concrete_method() {
    assert_eq!(run_prints(r#"<?php
abstract class Base {
    public function identify(): string { return "base"; }
    abstract public function tag(): string;
}
class Child extends Base {
    public function tag(): string { return "child"; }
}
$c = new Child();
echo $c->identify() . ',' . $c->tag();
"#), vec!["base,child"]);
}

// ── Abstract class with constructor ───────────────────────────

#[test] fn abstract_class_constructor_called_by_child() {
    assert_eq!(run_prints(r#"<?php
abstract class Vehicle {
    public function __construct(public readonly string $make) {}
    abstract public function type(): string;
}
class Car extends Vehicle {
    public function type(): string { return "car"; }
}
$c = new Car("Toyota");
echo $c->make . ':' . $c->type();
"#), vec!["Toyota:car"]);
}

// ── Abstract class with static method ────────────────────────

#[test] fn abstract_class_static_method_callable() {
    assert_eq!(run_prints(r#"<?php
abstract class Registry {
    private static array $items = [];
    public static function add(string $item): void { self::$items[] = $item; }
    public static function all(): array { return self::$items; }
}
class MyRegistry extends Registry {}
MyRegistry::add("a");
MyRegistry::add("b");
echo implode(',', MyRegistry::all());
"#), vec!["a,b"]);
}

// ── Abstract class chain ──────────────────────────────────────

#[test] fn abstract_class_two_levels_deep() {
    assert_eq!(run_prints(r#"<?php
abstract class A { abstract public function name(): string; }
abstract class B extends A { abstract public function age(): int; }
class C extends B {
    public function name(): string { return "Carol"; }
    public function age(): int { return 25; }
}
$c = new C();
echo $c->name() . ',' . $c->age();
"#), vec!["Carol,25"]);
}

// ── final class ──────────────────────────────────────────────

#[test] fn final_class_cannot_be_extended() {
    assert_eq!(run_prints(r#"<?php
final class Sealed {}
try {
    eval('class Child extends Sealed {}');
} catch (\Error $e) {
    echo "cannot extend";
}
"#), vec!["cannot extend"]);
}

#[test] fn final_class_can_be_instantiated() {
    assert_eq!(run_prints(r#"<?php
final class Config {
    public function __construct(public readonly string $env) {}
}
$c = new Config("production");
echo $c->env;
"#), vec!["production"]);
}

// ── final method ─────────────────────────────────────────────

#[test] fn final_method_cannot_be_overridden() {
    assert_eq!(run_prints(r#"<?php
class Base {
    final public function version(): string { return "1.0"; }
}
try {
    eval('class Child extends Base { public function version(): string { return "2.0"; } }');
} catch (\Error $e) {
    echo "cannot override";
}
"#), vec!["cannot override"]);
}

#[test] fn final_method_callable_on_child() {
    assert_eq!(run_prints(r#"<?php
class Parent2 {
    final public function id(): int { return 42; }
}
class Child2 extends Parent2 {}
echo (new Child2())->id();
"#), vec!["42"]);
}

// ── Abstract + interface combination ─────────────────────────

#[test] fn abstract_class_partially_implements_interface() {
    assert_eq!(run_prints(r#"<?php
interface Logger {
    public function log(string $msg): void;
    public function getLog(): array;
}
abstract class BaseLogger implements Logger {
    protected array $entries = [];
    public function getLog(): array { return $this->entries; }
}
class ConsoleLogger extends BaseLogger {
    public function log(string $msg): void { $this->entries[] = $msg; }
}
$l = new ConsoleLogger();
$l->log("hello");
$l->log("world");
echo implode(',', $l->getLog());
"#), vec!["hello,world"]);
}

// ── Abstract const (PHP 8.1) ──────────────────────────────────

#[test] fn abstract_class_with_class_constant() {
    assert_eq!(run_prints(r#"<?php
abstract class Protocol {
    const VERSION = '1.0';
    abstract public function connect(): void;
}
class HTTP extends Protocol {
    public function connect(): void { echo self::VERSION; }
}
(new HTTP())->connect();
"#), vec!["1.0"]);
}

// ── Template method pattern ───────────────────────────────────

#[test] fn template_method_pattern_with_abstract() {
    assert_eq!(run_prints(r#"<?php
abstract class Report {
    final public function generate(): string {
        return $this->header() . "\n" . $this->body() . "\n" . $this->footer();
    }
    abstract protected function header(): string;
    abstract protected function body(): string;
    protected function footer(): string { return "---end---"; }
}
class SalesReport extends Report {
    protected function header(): string { return "SALES REPORT"; }
    protected function body(): string { return "Total: $9999"; }
}
echo (new SalesReport())->generate();
"#), vec!["SALES REPORT\nTotal: $9999\n---end---"]);
}

// ── Abstract static methods ───────────────────────────────────

#[test] fn abstract_static_method_implemented_in_child() {
    assert_eq!(run_prints(r#"<?php
abstract class Container {
    abstract public static function type(): string;
    public static function describe(): string { return "Type: " . static::type(); }
}
class Box extends Container {
    public static function type(): string { return "box"; }
}
echo Box::describe();
"#), vec!["Type: box"]);
}

// ── Final on abstract mix — intermediate class ────────────────

#[test] fn abstract_method_final_in_intermediate_class() {
    assert_eq!(run_prints(r#"<?php
abstract class A2 { abstract public function run(): string; }
class B2 extends A2 { final public function run(): string { return "B2"; } }
class C2 extends B2 {}
echo (new C2())->run();
"#), vec!["B2"]);
}

// ── Multiple inheritance of abstract through interface chain ──

#[test] fn abstract_implements_interface_child_must_implement() {
    assert_eq!(run_prints(r#"<?php
interface Runnable { public function run(): void; }
abstract class Task implements Runnable {
    public function schedule(): void { echo "scheduled:"; $this->run(); }
}
class PrintTask extends Task {
    public function run(): void { echo "running"; }
}
(new PrintTask())->schedule();
"#), vec!["scheduled:running"]);
}

// ── Abstract class with property promotion ────────────────────

#[test] fn abstract_class_with_promoted_property() {
    assert_eq!(run_prints(r#"<?php
abstract class Entity {
    public function __construct(public readonly int $id) {}
    abstract public function label(): string;
}
class Product extends Entity {
    public function __construct(int $id, public readonly string $name) {
        parent::__construct($id);
    }
    public function label(): string { return "$this->id:$this->name"; }
}
echo (new Product(1, 'Widget'))->label();
"#), vec!["1:Widget"]);
}

// ── Abstract in trait ─────────────────────────────────────────

#[test] fn abstract_method_in_trait_forces_class_implementation() {
    assert_eq!(run_prints(r#"<?php
trait Validator {
    abstract protected function rules(): array;
    public function validate(array $data): bool {
        foreach ($this->rules() as $rule) if (!isset($data[$rule])) return false;
        return true;
    }
}
class Form {
    use Validator;
    protected function rules(): array { return ['email', 'password']; }
}
$f = new Form();
echo $f->validate(['email' => 'a@b.com', 'password' => 'x']) ? 'valid' : 'invalid';
"#), vec!["valid"]);
}

// ── Final prevents late static binding override ───────────────

#[test] fn final_class_late_static_binding_stays_in_class() {
    assert_eq!(run_prints(r#"<?php
final class Singleton {
    private static ?self $instance = null;
    private function __construct() {}
    public static function get(): static { return self::$instance ??= new self(); }
    public function whoAmI(): string { return static::class; }
}
echo Singleton::get()->whoAmI();
"#), vec!["Singleton"]);
}

// ── Abstract class with interface default + abstract hybrid ───

#[test] fn abstract_child_inherits_parent_concrete_methods() {
    assert_eq!(run_prints(r#"<?php
abstract class Writer {
    public function writeLine(string $s): void { echo $s . "\n"; }
    abstract public function target(): string;
}
abstract class NetworkWriter extends Writer {
    public function prefix(): string { return "[net] "; }
}
class HttpWriter extends NetworkWriter {
    public function target(): string { return "HTTP"; }
    public function write(string $msg): void { $this->writeLine($this->prefix() . $msg); }
}
(new HttpWriter())->write("hello");
"#), vec!["[net] hello"]);
}

// ── Abstract class with typed properties ─────────────────────

#[test] fn abstract_class_typed_property_accessible_in_child() {
    assert_eq!(run_prints(r#"<?php
abstract class DataSource {
    protected string $connection = '';
    abstract public function connect(string $dsn): void;
    public function getConnection(): string { return $this->connection; }
}
class DbSource extends DataSource {
    public function connect(string $dsn): void { $this->connection = $dsn; }
}
$db = new DbSource();
$db->connect("mysql://localhost/mydb");
echo $db->getConnection();
"#), vec!["mysql://localhost/mydb"]);
}

// ── instanceof works on abstract ──────────────────────────────

#[test] fn instanceof_abstract_parent_returns_true_for_child() {
    assert_eq!(run_prints(r#"<?php
abstract class Animal2 {}
class Cat extends Animal2 {}
$c = new Cat();
echo ($c instanceof Animal2) ? 'yes' : 'no';
"#), vec!["yes"]);
}

// ── Final and abstract cannot coexist ────────────────────────

#[test] fn compile_final_abstract_class_is_error() {
    assert_eq!(run_prints(r#"<?php
try {
    eval('abstract final class X {}');
} catch (\Error $e) {
    echo "error";
}
"#), vec!["error"]);
}
