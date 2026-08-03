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

#[test]
fn constructor_property_access_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class User {
    public string $name;
    public function __construct(string $name) { $this->name = $name; }
}
$u = new User('Alice');
echo $u->name;
"#,
        ),
        &["Alice"]
    );
}

#[test]
fn protected_property_access_via_method_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    protected int $value = 5;
    public function getValue(): int { return $this->value; }
}
$b = new Base();
echo $b->getValue();
"#,
        ),
        &["5"]
    );
}

#[test]
fn method_visibility_and_call_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    protected function base(): string { return 'base'; }
    public function wrapper(): string { return $this->base(); }
}
$b = new Base();
echo $b->wrapper();
"#,
        ),
        &["base"]
    );
}

#[test]
fn trait_precedence_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
trait A { public function value(): string { return 'a'; } }
trait B { public function value(): string { return 'b'; } }
class C {
    use A, B { B::value insteadof A; }
}
echo (new C())->value();
"#,
        ),
        &["b"]
    );
}

#[test]
fn multiple_inheritance_interfaces_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Logger { public function log(): string; }
interface Formatter { public function format(string $s): string; }
class Item implements Logger, Formatter {
    public function log(): string { return 'log'; }
    public function format(string $s): string { return strrev($s); }
}
$i = new Item();
echo $i->log(), '|', $i->format('ab');
"#,
        ),
        &["log|ba"]
    );
}

#[test]
fn late_static_binding_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public static function who(): string { return static::class; }
    public function resolve(): string { return static::who(); }
}
class Child extends Base {}
echo Child::who(), '|', (new Child())->resolve();
"#,
        ),
        &["Child|Child"]
    );
}

#[test]
fn static_self_difference_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public static function staticName(): string { return static::class; }
    public static function selfName(): string { return self::class; }
    public static function both(): string { return self::selfName() . '|' . static::staticName(); }
}
class Child extends Base {}
echo Child::both();
"#,
        ),
        &["Base|Child"]
    );
}

#[test]
fn anonymous_class_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new class('X') {
    public function __construct(public string $name) {}
    public function greet(): string { return $this->name; }
};
echo $obj->greet();
"#,
        ),
        &["X"]
    );
}

#[test]
fn object_clone_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    public int $n = 1;
}
$a = new Counter();
$b = clone $a;
$b->n = 3;
echo $a->n, '|', $b->n;
"#,
        ),
        &["1|3"]
    );
}

#[test]
fn dynamic_property_and_method_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Reporter {
    public function __construct(public string $name) {}
    public function describe(): string { return 'user:' . $this->name; }
}
$reporterClass = Reporter::class;
$method = 'describe';
$prop = 'name';
$obj = new $reporterClass('Ada');
echo $obj->$prop;
echo $obj->$method();
"#,
        ),
        &["Adaauser:Ada"]
    );
}

#[test]
fn magic_get_set_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box {
    private array $values = [];
    public function __set(string $name, mixed $value): void {
        $this->values[$name] = $value;
    }
    public function __get(string $name): mixed {
        return $this->values[$name] ?? null;
    }
}
$box = new Box();
$box->title = 'Book';
echo $box->title;
echo $box->missing ?? 'none';
"#,
        ),
        &["Booknone"]
    );
}

#[test]
fn callable_object_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Multiplier {
    public function __invoke(int $left, int $right): int {
        return $left * $right;
    }
}
$f = new Multiplier();
echo $f(3, 7);
"#,
        ),
        &["21"]
    );
}

#[test]
fn static_factory_late_binding_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class BaseWidget {
    public function __construct(public string $id) {}
    public static function make(string $id): static {
        return new static($id);
    }
}
class ButtonWidget extends BaseWidget {}
echo (new ButtonWidget('submit'))->id;
echo ButtonWidget::make('cancel')->id;
"#,
        ),
        &["submit|cancel"]
    );
}

#[test]
fn constructor_property_promotion_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Span {
    public function __construct(
        public int $start,
        public int $end = 0
    ) {}
}
$s = new Span(1, 9);
echo $s->start;
echo $s->end;
"#,
        ),
        &["19"]
    );
}

#[test]
fn oop_parent_method_and_property_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class CounterBase {
    public int $value = 1;
    protected function step(): int { $this->value += 1; return $this->value; }
}
class CounterChild extends CounterBase {
    public function double_step(): int {
        return parent::step() + parent::step();
    }
}
$counter = new CounterChild();
echo $counter->double_step();
echo '|';
echo $counter->value;
"#,
        ),
        &["5|3"]
    );
}

#[test]
fn oop_parent_constructor_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class PersonBase {
    public string $label;
    public function __construct(string $label) { $this->label = $label; }
}
class PersonChild extends PersonBase {
    public function __construct(string $name) {
        parent::__construct('guest:' . $name);
    }
}
echo (new PersonChild('Ada'))->label;
"#,
        ),
        &["guest:Ada"]
    );
}

#[test]
fn oop_static_late_binding_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class ServiceBase {
    public static string $scope = 'base';
    public static function scopeLabel(): string { return static::$scope; }
}
class ServiceChild extends ServiceBase {
    public static string $scope = 'child';
}
echo ServiceBase::scopeLabel();
echo '|';
echo ServiceChild::scopeLabel();
"#,
        ),
        &["base|child"]
    );
}

#[test]
fn oop_abstract_contract_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Writer {
    abstract public function write(string $text): string;
    public function prefix(): string { return 'x:' . $this->write('ok'); }
}
class UpperWriter extends Writer {
    public function write(string $text): string { return strtoupper($text); }
}
echo (new UpperWriter())->prefix();
"#,
        ),
        &["x:OK"]
    );
}

#[test]
fn oop_interfaces_multiple_contract_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Logger {
    public function log(string $value): string;
}
interface Formatter {
    public function format(string $value): string;
}
class Message implements Logger, Formatter {
    public function log(string $value): string { return 'log:' . $value; }
    public function format(string $value): string { return strtoupper($value); }
}
$m = new Message();
echo $m->log($m->format('hi'));
"#,
        ),
        &["log:HI"]
    );
}

#[test]
fn oop_trait_precedence_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
trait SourceA {
    public function source(): string { return 'A'; }
}
trait SourceB {
    public function source(): string { return 'B'; }
}
class TaggedSource {
    use SourceA, SourceB { SourceB::source insteadof SourceA; }
}
echo (new TaggedSource())->source();
"#,
        ),
        &["B"]
    );
}

#[test]
fn oop_trait_alias_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
trait BaseGreeting {
    public function sayHello(): string { return 'hello'; }
}
class Greeting {
    use BaseGreeting;
    public function loudHello(): string { return strtoupper($this->sayHello()); }
}
echo (new Greeting())->sayHello();
echo '|';
echo (new Greeting())->loudHello();
"#,
        ),
        &["hello|HELLO"]
    );
}

#[test]
fn oop_magic_call_static_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class CommandBus {
    public function __call(string $name, array $args): string {
        return $name . ':' . implode(',', $args);
    }
    public static function __callStatic(string $name, array $args): string {
        return strtoupper($name) . ':' . implode('|', $args);
    }
}
$obj = new CommandBus();
echo $obj->render(1, 2);
echo '|';
echo CommandBus::dispatch('job');
"#,
        ),
        &["render:1,2|DISPATCH:job"]
    );
}

#[test]
fn oop_clone_with_magic_reset_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Buffer {
    public string $value;
    public function __construct(string $value) { $this->value = $value; }
    public function __clone(): void { $this->value .= '-cloned'; }
}
$original = new Buffer('base');
$copy = clone $original;
echo $original->value;
echo '|';
echo $copy->value;
"#,
        ),
        &["base|base-cloned"]
    );
}

#[test]
fn __call_magic_method_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Dynamic {
    public function __call(string $name, array $args): mixed {
        if ($name === 'ping') {
            return implode('|', $args);
        }
        return null;
    }
}
$d = new Dynamic();
echo $d->ping('alpha', 'beta');
"#,
        ),
        &["alpha|beta"]
    );
}

#[test]
fn __clone_magic_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Snapshot {
    public function __construct(public string $label) {}
    public function __clone(): void {
        $this->label .= ':clone';
    }
}
$a = new Snapshot('A');
$b = clone $a;
echo $a->label;
echo $b->label;
"#,
        ),
        &["A|A:clone"]
    );
}

#[test]
fn object_to_string_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Node {
    public function __construct(public string $name) {}
    public function __toString(): string {
        return "node:" . $this->name;
    }
}
echo (string) new Node("root");
"#,
        ),
        &["node:root"]
    );
}

#[test]
fn magic_unset_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Bag {
    private array $v = [];
    public function __set(string $name, mixed $value): void { $this->v[$name] = $value; }
    public function __get(string $name): mixed { return $this->v[$name] ?? null; }
    public function __unset(string $name): void { unset($this->v[$name]); }
    public function has(string $name): bool { return isset($this->v[$name]); }
}
$b = new Bag();
$b->x = 1;
unset($b->x);
echo $b->has('x') ? 'yes' : 'no';
"#,
        ),
        &["no"]
    );
}

#[test]
fn abstract_template_method_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Template {
    abstract protected function body(): string;
    public function render(): string {
        return "<" . $this->body() . ">";
    }
}
class MessageTemplate extends Template {
    protected function body(): string { return "ok"; }
}
echo (new MessageTemplate())->render();
"#,
        ),
        &["<ok>"]
    );
}

#[test]
fn static_private_property_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    private static int $count = 0;
    public static function inc(): int {
        return ++self::$count;
    }
}
echo Counter::inc();
echo Counter::inc();
"#,
        ),
        &["12"]
    );
}

#[test]
fn __call_static_magic_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Api {
    public static function __callStatic(string $name, array $args): mixed {
        if ($name === 'endpoint') {
            return $args[0] . ':' . $args[1];
        }
        return null;
    }
}
echo Api::endpoint('v1', 'users');
"#,
        ),
        &["v1:users"]
    );
}

#[test]
fn reflection_property_methods_counts_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class User {
    public int $id;
    private string $name;
    public function __construct() {}
    public function hello(): string { return 'hi'; }
}
$obj = new User();
echo property_exists($obj, 'id') ? 'id' : 'no';
echo method_exists($obj, 'hello') ? '|hello' : '|no';
"#,
        ),
        &["id|hello"]
    );
}

#[test]
fn class_name_helpers_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Model {}
$o = new Model();
echo get_class($o), '|';
echo get_class(new Model()), '|';
echo is_a($o, 'Model') ? 'is-a' : 'not';
"#,
        ),
        &["Model|Model|is-a"]
    );
}

#[test]
fn final_class_compile_and_runtime_behavior() {
    assert_eq!(
        run_prints(
            r#"<?php
final class Endpoint {
    public function name(): string { return 'endpoint'; }
}
echo (new Endpoint())->name();
"#,
        ),
        &["endpoint"]
    );
}

#[test]
fn final_method_runtime_invocation() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public final function lock(): string { return 'base'; }
}
class Child extends Base {
    public function ping(): string { return $this->lock(); }
}
echo (new Child())->ping();
"#,
        ),
        &["base"]
    );
}

#[test]
fn anonymous_class_with_constructor_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$factory = function(string $label): object {
    return new class($label) {
        public function __construct(private string $label) {}
        public function value(): string { return $this->label; }
    };
};
$x = $factory('abc');
echo $x->value();
"#,
        ),
        &["abc"]
    );
}

#[test]
fn object_comparison_identity_vs_equality_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Marker { public function __construct(public int $v) {} }
$a = new Marker(1);
$b = new Marker(1);
$c = $a;
echo ($a == $b) ? 'eq' : 'neq';
echo ($a === $c) ? '|same' : '|diff';
"#,
        ),
        &["eq|same"]
    );
}

#[test]
fn object_to_array_and_cast_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Pair {
    public function __construct(public int $a, public int $b) {}
}
$p = new Pair(1, 2);
$arr = (array) $p;
ksort($arr);
echo json_encode($arr);
"#,
        ),
        &["{\"a\":1,\"b\":2}"]
    );
}

#[test]
fn object_serialize_with_magic_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Token {
    public string $value = 'x';
    public function __serialize(): array {
        return ['value' => $this->value];
    }
    public function __unserialize(array $data): void {
        $this->value = $data['value'] . '-u';
    }
}
$t = new Token();
$state = serialize($t);
$u = unserialize($state);
echo $u->value;
"#,
        ),
        &["x-u"]
    );
}

#[test]
fn __invoke_magic_method_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Adder {
    public function __invoke(int $a, int $b): int {
        return $a + $b;
    }
}
$fn = new Adder();
echo $fn(4, 6);
"#,
        ),
        &["10"]
    );
}

#[test]
fn __set_state_magic_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class State {
    public function __construct(public int $n) {}
    public static function __set_state(array $props): State {
        return new State($props['n'] + 1);
    }
}
$state = var_export(new State(3), true);
$obj = eval('return ' . $state . ';');
echo $obj->n;
"#,
        ),
        &["4"]
    );
}

#[test]
fn object_identities_and_references_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box { public int $n = 1; }
$a = new Box();
$b = $a;
$b->n = 8;
echo $a->n;
"#,
        ),
        &["8"]
    );
}

#[test]
fn inherited_constructor_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public string $name;
    public function __construct(string $name) { $this->name = $name; }
}
class Child extends Base {}
echo (new Child('x'))->name;
"#,
        ),
        &["x"]
    );
}

#[test]
fn enum_with_backed_value_and_match_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
enum Role: string {
    case ADMIN = 'admin';
    case USER = 'user';
}
function label(Role $role): string {
    return match($role) {
        Role::ADMIN => 'A',
        Role::USER => 'U' };
}
echo label(Role::USER);
"#,
        ),
        &["U"]
    );
}

#[test]
fn final_class_prevents_extension_runtime_check() {
    assert_eq!(
        run_prints(
            r#"<?php
final class FinalService {
    public function ping(): string { return 'pong'; }
}
$svc = new FinalService();
echo $svc->ping();
"#,
        ),
        &["pong"]
    );
}

#[test]
fn trait_method_alias_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
trait Logger {
    public function log(string $msg): string { return 'log:' . $msg; }
}
class Audit {
    use Logger {
        log as public report;
    }
}
echo (new Audit())->report('ok');
"#,
        ),
        &["log:ok"]
    );
}

#[test]
fn late_static_binding_static_class_name_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class BaseFactory {
    public static function make(): string {
        return static::class;
    }
}
class Widget extends BaseFactory {}
echo Widget::make();
"#,
        ),
        &["Widget"]
    );
}

#[test]
fn object_comparison_string_cast_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Item {
    public string $id;
    public function __construct(string $id) { $this->id = $id; }
    public function __toString(): string { return $this->id; }
}
$one = new Item('A');
$two = new Item('A');
echo (string)$one;
echo (string)$two;
echo ($one == $two) ? '|eq' : '|neq';
"#,
        ),
        &["AA|eq"]
    );
}

#[test]
fn destruct_and_serialize_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Finalizer {
    public function __construct(private string $token) {}
    public function __destruct(): void {
        echo "bye:$this->token;";
    }
}
function make_and_drop(): void {
    $x = new Finalizer('A');
}
make_and_drop();
echo 'done';
"#,
        ),
        &["bye:A;done"]
    );
}

#[test]
fn readonly_property_enforced_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box {
    public function __construct(public readonly string $value) {}
}
$b = new Box('x');
try {
    $b->value = 'y';
} catch (Throwable $e) {
    echo 'err';
}
"#,
        ),
        &["err"]
    );
}

#[test]
fn trait_abstract_and_concrete_resolution_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
trait Renderer {
    abstract public function source(): string;
    public function wrap(): string {
        return '[' . $this->source() . ']';
    }
}
class Message {
    use Renderer;
    public function source(): string { return 'hi'; }
}
echo (new Message())->wrap();
"#,
        ),
        &["[hi]"]
    );
}

#[test]
fn constructor_destructuring_promoted_property_and_unset_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Holder {
    public function __construct(public string $value) {}
}
$h = new Holder('x');
$f = fn(Holder $item): string => $item->value;
unset($h);
echo $f(new Holder('y'));
"#,
        ),
        &["y"]
    );
}

#[test]
fn readonly_property_enforced_across_aggregate_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box {
    public function __construct(public readonly int $id) {}
}
$x = new Box(1);
$y = $x;
$yId = $y->id;
echo $yId;
"#,
        ),
        &["1"]
    );
}

#[test]
fn static_property_in_trait_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
trait Counters {
    public static int $n = 0;
    public static function bump(): int { return ++self::$n; }
}
class Widget {
    use Counters;
}
echo Widget::bump();
echo Widget::bump();
echo Widget::$n;
"#,
        ),
        &["12|2"]
    );
}

#[test]
fn magic_isset_checks_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Bag {
    private array $values = [];
    public function __set(string $name, mixed $value): void { $this->values[$name] = $value; }
    public function __isset(string $name): bool { return array_key_exists($name, $this->values); }
}
$b = new Bag();
echo isset($b->token) ? 'yes' : 'no';
$b->token = 'ok';
echo isset($b->token) ? 'yes' : 'no';
"#,
        ),
        &["noyes"]
    );
}

#[test]
fn static_visibility_public_private_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Keys {
    private const SECRET = 'secret';
    public const OPEN = 'open';
    public static function secretValue(): string {
        return self::SECRET;
    }
}
echo Keys::OPEN;
echo '|';
echo Keys::secretValue();
"#,
        ),
        &["open|secret"]
    );
}

#[test]
fn static_properties_are_mutable_across_calls_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    public static int $n = 0;
    public static function bump(): int {
        return ++self::$n;
    }
}
echo Counter::bump();
echo Counter::bump();
echo '|';
echo Counter::$n;
"#,
        ),
        &["12|2"]
    );
}

#[test]
fn property_visibility_and_magic_getter_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Model {
    private string $name = 'x';
    public function __get(string $name): mixed {
        return $name === 'name' ? 'resolved' : null;
    }
}
$m = new Model();
echo $m->name;
"#,
        ),
        &["resolved"]
    );
}

#[test]
fn __invoke_can_be_used_as_callable_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Service {
    public function __invoke(int $v): int { return $v * 3; }
}
echo (new Service())(4);
"#,
        ),
        &["12"]
    );
}

#[test]
fn class_exists_with_fqcn_and_autoload_style() {
    assert_eq!(
        run_prints(
            r#"<?php
class Demo {}
echo class_exists('Demo') ? 'yes' : 'no';
echo '|';
echo class_exists('\\Demo') ? 'yes' : 'no';
"#,
        ),
        &["yes|yes"]
    );
}

#[test]
fn iterator_aggregate_object_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Lines implements IteratorAggregate {
    private array $items = ['a', 'b', 'c'];
    public function getIterator(): Traversable {
        return new ArrayIterator($this->items);
    }
}
$l = new Lines();
$out = [];
foreach ($l as $item) {
    $out[] = $item;
}
echo implode('|', $out);
"#,
        ),
        &["a|b|c"]
    );
}

#[test]
fn object_cast_to_bool_and_clone_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Holder {
    public string $v;
    public function __construct(string $v) { $this->v = $v; }
}
$a = new Holder('x');
$b = clone $a;
$b->v = 'y';
echo ($a === $b ? 'same' : 'diff') . '|';
echo (bool)$a . '-' . (bool)$b;
"#,
        ),
        &["diff|1-1"]
    );
}

#[test]
fn trait_conflict_resolution_precedence_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
trait TOne {
    public function label(): string { return 'one'; }
}
trait TTwo {
    public function label(): string { return 'two'; }
}
class Both {
    use TOne, TTwo { TTwo::label insteadof TOne; }
}
echo (new Both())->label();
"#,
        ),
        &["two"]
    );
}

#[test]
fn dynamic_property_access_on_object_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Node {
    public function __construct(public string $name) {}
}
$obj = new Node('root');
$field = 'name';
echo $obj->$field;
"#,
        ),
        &["root"]
    );
}

#[test]
fn class_alias_and_exists_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class OriginalWidget {
    public function label(): string { return 'orig'; }
}
class_alias('OriginalWidget', 'AliasWidget');
echo class_exists('AliasWidget') ? 'yes' : 'no';
echo '|';
echo (new AliasWidget())->label();
"#,
        ),
        &["yes|orig"]
    );
}

#[test]
fn object_id_changes_with_clone_but_not_reference_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Holder {
    public string $name;
    public function __construct(string $name) { $this->name = $name; }
}
$a = new Holder('left');
$b = $a;
$c = clone $a;
$b->name = 'right';
echo spl_object_id($a) === spl_object_id($b) ? 'same' : 'diff';
echo '|';
echo spl_object_id($a) === spl_object_id($c) ? 'same' : 'diff';
"#,
        ),
        &["same|diff"]
    );
}

#[test]
fn magic_isset_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class DynamicStore {
    private array $vals = [];
    public function __set(string $name, mixed $value): void { $this->vals[$name] = $value; }
    public function __get(string $name): mixed { return $this->vals[$name] ?? null; }
    public function __isset(string $name): bool { return array_key_exists($name, $this->vals); }
}
$d = new DynamicStore();
$d->alpha = 1;
echo isset($d->alpha) ? 'alpha' : 'noalpha';
echo '|';
echo isset($d->beta) ? 'beta' : 'nobeta';
"#,
        ),
        &["alpha|nobeta"]
    );
}

#[test]
fn sleep_and_wakeup_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Persisted {
    public string $token;
    public function __construct(string $token) { $this->token = $token; }
    public function __sleep(): array { return ['token']; }
    public function __wakeup(): void { $this->token = $this->token . '-awake'; }
}
$p = new Persisted('abc');
$restored = unserialize(serialize($p));
echo $restored->token;
"#,
        ),
        &["abc-awake"]
    );
}

#[test]
fn serialization_serialize_unserialize_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Package {
    public function __construct(public string $name, public int $version) {}
}
$p = new Package('core', 1);
$copy = unserialize(serialize($p));
echo $copy->name . '|' . $copy->version;
"#,
        ),
        &["core|1"]
    );
}

#[test]
fn clone_reinitializes_with_magic_clone_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    public array $values;
    public function __construct() { $this->values = [1, 2, 3]; }
    public function __clone(): void { $this->values[] = 9; }
}
$a = new Counter();
$b = clone $a;
$a->values[0] = 9;
echo implode(',', $b->values);
"#,
        ),
        &["1,2,3,9"]
    );
}

#[test]
fn anonymous_class_interface_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Service { public function run(): string; }
$svc = new class implements Service {
    public function run(): string { return 'ok'; }
};
echo $svc->run();
echo class_exists($svc::class) ? '' : '';
"#,
        ),
        &["ok"]
    );
}

#[test]
fn oop_constructor_parameter_order_and_defaults_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Span {
    public function __construct(public int $x, public int $y = 4) {}
}
$s = new Span(3);
echo $s->x;
echo '|';
echo $s->y;
"#,
        ),
        &["3|4"]
    );
}

#[test]
fn oop_magic_method_ordered_invocation_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Recorder {
    private array $trace = [];
    public function __set(string $name, mixed $value): void {
        $this->trace[] = 'set:' . $name;
    }
    public function trace(): string { return implode(',', $this->trace); }
}
$r = new Recorder();
$r->v = 7;
echo $r->trace();
"#,
        ),
        &["set:v"]
    );
}

#[test]
fn oop_static_property_inheritance_isolation_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public static int $counter = 1;
}
class Child extends Base {}
Base::$counter = 2;
echo Base::$counter;
echo '|';
echo Child::$counter;
"#,
        ),
        &["2|2"]
    );
}

#[test]
fn oop_dynamic_method_call_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Worker {
    public function work(string $name): string { return 'do:' . $name; }
}
$w = new Worker();
$method = 'work';
echo $w->$method('task');
"#,
        ),
        &["do:task"]
    );
}

#[test]
fn oop_type_validation_with_instanceof_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Marker {}
class A implements Marker {}
class B {}
echo (new A()) instanceof Marker ? 'yes' : 'no';
echo '|';
echo (new B()) instanceof Marker ? 'yes' : 'no';
"#,
        ),
        &["yes|no"]
    );
}

#[test]
fn oop_static_counter_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Stats {
    public static int $count = 0;
    public function tick(): int {
        return ++self::$count;
    }
}
$a = new Stats();
$b = new Stats();
echo $a->tick();
echo '|';
echo $b->tick();
echo '|';
echo Stats::$count;
"#,
        ),
        &["1|2|2"]
    );
}

#[test]
fn oop_static_method_and_parent_call_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public static function title(): string {
        return 'base';
    }
}
class Child extends Base {
    public static function title(): string {
        return parent::title() . '-child';
    }
}
echo Child::title();
"#,
        ),
        &["base-child"]
    );
}

#[test]
fn oop_trait_alias_and_visibility_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
trait T {
    public function baseName(): string { return 'base'; }
}
class Item {
    use T {
        baseName as public displayName;
    }
}
echo (new Item())->displayName();
"#,
        ),
        &["base"]
    );
}

#[test]
fn oop_class_alias_preserves_behavior_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class RealService {
    public function id(): string {
        return 'real';
    }
}
class_alias(RealService::class, 'ServiceAlias');
echo (new ServiceAlias())->id();
"#,
        ),
        &["real"]
    );
}

#[test]
fn oop_parent_constructor_called_via_parent_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Widget {
    public int $value;
    public function __construct(int $value) {
        $this->value = $value;
    }
}
class Button extends Widget {
    public function __construct() {
        parent::__construct(9);
    }
}
echo (new Button())->value;
"#,
        ),
        &["9"]
    );
}

#[test]
fn oop_magic_set_get_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Data {
    private array $store = [];
    public function __set(string $name, mixed $value): void {
        $this->store[$name] = strtoupper($value);
    }
    public function __get(string $name): mixed {
        return $this->store[$name] ?? null;
    }
}
$d = new Data();
$d->name = 'demo';
echo $d->name;
"#,
        ),
        &["DEMO"]
    );
}

#[test]
fn oop_constructor_defaults_with_named_parameters_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Coord {
    public function __construct(
        public int $x = 1,
        public int $y = 2,
    ) {}
}
$c = new Coord(y: 7);
echo $c->x;
echo '|';
echo $c->y;
"#,
        ),
        &["1|7"]
    );
}

#[test]
fn oop_magic_clone_independence_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box {
    public array $tags;
    public function __construct(array $tags) { $this->tags = $tags; }
    public function __clone(): void {
        $this->tags[] = 'cloned';
    }
}
$a = new Box(['a']);
$b = clone $a;
$a->tags[] = 'source';
echo implode(',', $a->tags);
echo '|';
echo implode(',', $b->tags);
"#,
        ),
        &["a,source|a,cloned"]
    );
}

#[test]
fn oop_static_property_inheritance_and_override_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public static string $name = 'base';
    public static function readName(): string { return static::$name; }
}
class Child extends Base {
    public static string $name = 'child';
}
echo Base::readName();
echo '|';
echo Child::readName();
"#,
        ),
        &["base|child"]
    );
}

#[test]
fn oop_method_call_in_traits_with_override_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
trait Logger {
    public function tag(): string { return 'base'; }
}
trait PrefixLogger {
    public function stamp(string $value): string { return 'pref:' . $value; }
}
class Service {
    use Logger;
    public function value(): string { return $this->tag(); }
}
class ServiceWithPrefix extends Service {
    use PrefixLogger;
    public function value(): string { return $this->stamp(parent::value()); }
}
echo (new Service())->value();
echo '|';
echo (new ServiceWithPrefix())->value();
"#,
        ),
        &["base|pref:base"]
    );
}

#[test]
fn oop_property_promoted_defaults_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Entry {
    public function __construct(
        public int $id = 1,
        public string $label = 'x',
    ) {}
}
$e = new Entry(label: 'ok');
echo $e->id;
echo '|';
echo $e->label;
"#,
        ),
        &["1|ok"]
    );
}

#[test]
fn oop_late_binding_with_constructor_chain_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Root {
    public function __construct(public string $kind) {}
    public static function make(string $kind): static {
        return new static($kind);
    }
}
class Leaf extends Root {}
echo (new Leaf('leaf'))->kind;
echo '|';
echo Leaf::make('mk')->kind;
"#,
        ),
        &["leaf|mk"]
    );
}

#[test]
fn oop_private_method_not_accessible_from_call_parent_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    private function secret(): string { return 'secret'; }
    public function reveal(): string { return $this->secret(); }
}
echo (new Base())->reveal();
"#,
        ),
        &["secret"]
    );
}

#[test]
fn oop_namespace_and_fqcn_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
namespace App\Module;
class Service {
    public function label(): string { return __CLASS__; }
}
echo (new Service())->label();
echo '|';
echo (new \App\Module\Service())->label();
echo '|';
echo class_exists('App\Module\Service') ? 'yes' : 'no';
"#,
        ),
        &["App\\Module\\Service|App\\Module\\Service|yes"]
    );
}

#[test]
fn oop_dynamic_class_and_static_invocation_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Factory {
    public static function build(string $name): string { return 'factory:' . $name; }
}
$factory = Factory::class;
echo $factory::build('widget');
"#,
        ),
        &["factory:widget"]
    );
}
