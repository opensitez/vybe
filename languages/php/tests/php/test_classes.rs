use super::helpers::{compile_ok, run_prints};

#[test]
fn class_empty() {
    compile_ok("<?php class Foo {} $f = new Foo();");
}
#[test]
fn class_constructor() {
    compile_ok(
        "<?php class Dog { public $name; public function __construct($name) { $this->name = $name; } } $d = new Dog('Rex');",
    );
}
#[test]
fn class_method() {
    compile_ok(
        "<?php class Dog { public $name; public function __construct($n) { $this->name = $n; } public function speak() { return $this->name . ' says Woof'; } } $d = new Dog('Rex'); echo $d->speak();",
    );
}
#[test]
fn class_inheritance() {
    compile_ok(
        "<?php class Animal { public $name; public function __construct($n) { $this->name = $n; } } class Cat extends Animal { public function speak() { return $this->name . ' says Meow'; } } $c = new Cat('Whiskers'); echo $c->speak();",
    );
}
#[test]
fn class_property_default() {
    compile_ok(
        "<?php class Config { public $debug = false; public $version = '1.0'; } $c = new Config();",
    );
}
#[test]
fn static_method() {
    compile_ok(
        "<?php class MathHelper { public static function square($n) { return $n * $n; } } echo MathHelper::square(5);",
    );
}
#[test]
fn class_constant() {
    compile_ok("<?php class Config { const VERSION = '1.0'; } echo Config::VERSION;");
}
#[test]
fn multiple_methods() {
    compile_ok(
        "<?php class Calc { public function add($a,$b) { return $a+$b; } public function sub($a,$b) { return $a-$b; } } $c = new Calc(); echo $c->add(3,2);",
    );
}
#[test]
fn chained_calls() {
    compile_ok(
        "<?php class Builder { public $val = ''; public function add($s) { $this->val .= $s; return $this; } } $b = new Builder(); $b->add('a')->add('b');",
    );
}
#[test]
fn new_with_args() {
    compile_ok(
        "<?php class Point { public $x; public $y; public function __construct($x, $y) { $this->x = $x; $this->y = $y; } } $p = new Point(1, 2);",
    );
}

#[test]
fn class_constructor_reads_name() {
    let out = run_prints(
        "<?php\nclass Dog { public string $name; public function __construct(string $name) { $this->name = $name; } }\n$d = new Dog('Rex');\necho $d->name;\n",
    );
    assert_eq!(out, vec!["Rex"]);
}

#[test]
fn class_method_returns_value() {
    let out = run_prints(
        "<?php\nclass Dog {\n    public function speak(string $name): string { return $name . ' says woof'; }\n}\n$d = new Dog();\necho $d->speak('Ada');\n",
    );
    assert_eq!(out, vec!["Ada says woof"]);
}

#[test]
fn class_inheritance_and_override() {
    let out = run_prints(
        "<?php\nclass Vehicle { public function noise(): string { return 'v'; } }\nclass Car extends Vehicle { public function noise(): string { return 'c'; } }\n$c = new Car();\necho $c->noise();\n",
    );
    assert_eq!(out, vec!["c"]);
}

#[test]
fn class_property_mutation() {
    let out = run_prints(
        "<?php\nclass Counter { public int $n = 0; public function inc(): void { $this->n += 1; } }\n$c = new Counter();\n$c->inc();\n$c->inc();\necho $c->n;\n",
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn class_static_property_and_method() {
    let out = run_prints(
        "<?php\nclass Math {\n    public static int $base = 2;\n    public static function scale(int $v): int { return self::$base * $v; }\n}\necho Math::scale(3), \"\\n\", Math::$base;\n",
    );
    assert_eq!(out, vec!["6", "2"]);
}

#[test]
fn class_constant_access() {
    let out = run_prints(
        "<?php\nclass Config { public const VERSION = '2.1'; }\necho Config::VERSION;\n",
    );
    assert_eq!(out, vec!["2.1"]);
}

#[test]
fn class_instanceof_check() {
    let out = run_prints(
        "<?php\ninterface Renderable {}\nclass Button implements Renderable {}\necho (new Button()) instanceof Renderable ? 'yes' : 'no';\n",
    );
    assert_eq!(out, vec!["yes"]);
}

#[test]
fn class_this_chain() {
    let out = run_prints(
        "<?php\nclass Builder {\n    public string $value = '';\n    public function append(string $s): self { $this->value .= $s; return $this; }\n}\necho (new Builder())->append('a')->append('b')->value;\n",
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn class_copy_by_reference_with_cloning() {
    let out = run_prints(
        "<?php\nclass Box { public string $label; public function __construct(string $label) { $this->label = $label; } }\n$b1 = new Box('x');\n$b2 = clone $b1;\n$b2->label = 'y';\necho $b1->label, '|', $b2->label;\n",
    );
    assert_eq!(out, vec!["x|y"]);
}

#[test]
fn class_constructor_property_access() {
    let out = run_prints(
        "<?php\nclass Point { public function __construct(public int $x, public int $y) {} }\n$p = new Point(1, 2);\necho $p->x + $p->y;\n",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn class_magic_methods_to_string() {
    let out = run_prints(
        "<?php\nclass Tag {\n    public function __construct(public string $name) {}\n    public function __toString(): string { return strtoupper($this->name); }\n}\necho new Tag('php');\n",
    );
    assert_eq!(out, vec!["PHP"]);
}

#[test]
fn class_private_property_isolation() {
    let out = run_prints(
        "<?php\nclass User {\n    private string $secret = 'x';\n    public function reveal(): string { return $this->secret; }\n}\n$u = new User();\necho $u->reveal();\n",
    );
    assert_eq!(out, vec!["x"]);
}

#[test]
fn class_final_method_in_base_invocation() {
    let out = run_prints(
        "<?php\nclass Base {\n    final public function id(): string { return 'base'; }\n    public function label(): string { return $this->id(); }\n}\nclass Child extends Base {}\necho (new Child())->label();\n",
    );
    assert_eq!(out, vec!["base"]);
}

#[test]
fn class_destructor_runs() {
    let out = run_prints(
        "<?php\nclass Session {\n    public function __destruct() { echo 'closed'; }\n}\n$s = new Session();\necho 'done';\nunset($s);\n",
    );
    assert_eq!(out, vec!["doneclosed"]);
}

#[test]
fn class_abstract_and_implementation_runtime() {
    let out = run_prints(
        "<?php
abstract class Provider {
    abstract public function id(): string;
    public function label(): string { return 'provider:' . $this->id(); }
}
class UserProvider extends Provider {
    public function id(): string { return 'user'; }
}
echo (new UserProvider())->label();
",
    );
    assert_eq!(out, vec!["provider:user"]);
}

#[test]
fn class_interface_method_contract_runtime() {
    let out = run_prints(
        "<?php
interface Renderable {
    public function render(): string;
}
class Card implements Renderable {
    public function render(): string { return 'card'; }
}
echo (new Card())->render();
",
    );
    assert_eq!(out, vec!["card"]);
}

#[test]
fn class_trait_with_method_alias_runtime() {
    let out = run_prints(
        "<?php
trait Logger {
    public function message(): string { return 'base'; }
}
class Service {
    use Logger {
        message as public aliasMessage;
    }
}
echo (new Service())->aliasMessage();
",
    );
    assert_eq!(out, vec!["base"]);
}

#[test]
fn class_trait_precedence_and_conflict_runtime() {
    let out = run_prints(
        "<?php
trait A {
    public function value(): string { return 'A'; }
}
trait B {
    public function value(): string { return 'B'; }
}
class Combined {
    use A, B {
        B::value insteadof A;
    }
}
echo (new Combined())->value();
",
    );
    assert_eq!(out, vec!["B"]);
}

#[test]
fn class_static_binding_get_called_class_runtime() {
    let out = run_prints(
        "<?php
class Base {
    public static function name(): string {
        return static::class;
    }
}
class Child extends Base {}
echo Base::name();
echo '|';
echo Child::name();
",
    );
    assert_eq!(out, vec!["Base|Child"]);
}

#[test]
fn class_property_visibility_runtime() {
    let out = run_prints(
        "<?php
class Model {
    private string $secret = 's';
    protected string $state = 'ok';
    public string $public = 'pub';
    public function readSecret(): string { return $this->secret; }
    public function readState(): string { return $this->state; }
}
$m = new Model();
echo $m->public;
echo '|';
echo $m->readSecret();
echo '|';
echo $m->readState();
",
    );
    assert_eq!(out, vec!["pub|s|ok"]);
}

#[test]
fn class_constructor_property_promotion_runtime() {
    let out = run_prints(
        "<?php
class Point {
    public function __construct(public int $x, public int $y) {}
}
$p = new Point(3, 4);
echo $p->x;
echo '|';
echo $p->y;
",
    );
    assert_eq!(out, vec!["3|4"]);
}

#[test]
fn class_magic_setter_getter_runtime() {
    let out = run_prints(
        "<?php
class Bag {
    private array $values = [];
    public function __set(string $name, mixed $value): void { $this->values[$name] = $value; }
    public function __get(string $name): mixed { return $this->values[$name] ?? null; }
}
$b = new Bag();
$b->lang = 'php';
echo $b->lang;
",
    );
    assert_eq!(out, vec!["php"]);
}

#[test]
fn class_readonly_property_initialization_runtime() {
    let out = run_prints(
        "<?php
class User {
    public function __construct(public readonly string $name) {}
}
$u = new User('alice');
echo $u->name;
",
    );
    assert_eq!(out, vec!["alice"]);
}

#[test]
fn class_magic_invoke_runtime() {
    let out = run_prints(
        "<?php
class InvokableCounter {
    private int $count = 0;
    public function __invoke(int $delta): int {
        $this->count += $delta;
        return $this->count;
    }
}
$svc = new InvokableCounter();
echo $svc(3);
echo '|';
echo $svc(2);
",
    );
    assert_eq!(out, vec!["3|5"]);
}

#[test]
fn class_magic_isset_unset_runtime() {
    let out = run_prints(
        "<?php
class Bucket {
    private array $values = [];
    public function __set(string $name, mixed $value): void { $this->values[$name] = $value; }
    public function __get(string $name): mixed { return $this->values[$name] ?? null; }
    public function __isset(string $name): bool { return array_key_exists($name, $this->values); }
    public function __unset(string $name): void { unset($this->values[$name]); }
}
$b = new Bucket();
$b->lang = 'php';
echo isset($b->lang) ? 'yes' : 'no';
echo '|';
unset($b->lang);
echo isset($b->lang) ? 'yes' : 'no';
",
    );
    assert_eq!(out, vec!["yes|no"]);
}

#[test]
fn class_clone_with_custom_clone_runtime() {
    let out = run_prints(
        "<?php
class CopyState {
    public int $hits = 0;
    public function __clone(): void { $this->hits = 99; }
}
$a = new CopyState();
$a->hits = 1;
$b = clone $a;
echo $b->hits;
",
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn class_final_class_compilation_only_runtime() {
    compile_ok(
        "<?php
final class Sealed {}
new Sealed();
",
    );
}

#[test]
fn class_static_late_binding_method_override_runtime() {
    let out = run_prints(
        "<?php
class Base {
    public static function factory(): static {
        return new static();
    }
}
class Child extends Base {}
echo (new Base())::factory() instanceof Base ? 'base' : 'not';
echo '|';
echo Child::factory() instanceof Child ? 'child' : 'no';
",
    );
    assert_eq!(out, vec!["base|child"]);
}

#[test]
fn class_implements_multiple_interfaces_runtime() {
    let out = run_prints(
        "<?php
interface A { public function a(): string; }
interface B { public function b(): string; }
class Impl implements A, B {
    public function a(): string { return 'A'; }
    public function b(): string { return 'B'; }
}
echo (new Impl()) instanceof A ? 'A' : 'no';
echo '|';
echo (new Impl()) instanceof B ? 'B' : 'no';
",
    );
    assert_eq!(out, vec!["A|B"]);
}

#[test]
fn class_instanceof_self_parent_runtime() {
    let out = run_prints(
        "<?php
class Vehicle { public function isVehicle(self $v): bool { return $v instanceof self; } }
class Car extends Vehicle {}
echo (new Vehicle())->isVehicle(new Vehicle()) ? 'yes' : 'no';
echo '|';
echo (new Vehicle())->isVehicle(new Car()) ? 'yes' : 'no';
",
    );
    assert_eq!(out, vec!["yes|no"]);
}

#[test]
fn class_static_property_in_subclass_runtime() {
    let out = run_prints(
        "<?php
class Base {
    public static int $limit = 1;
}
class Child extends Base {
    public static int $limit = 2;
}
echo Base::$limit;
echo '|';
echo Child::$limit;
",
    );
    assert_eq!(out, vec!["1|2"]);
}

#[test]
fn class_abstract_trait_requirement_runtime() {
    let out = run_prints(
        "<?php
trait HasFormatter {
    abstract public function format(): string;
    public function output(): string { return '[' . $this->format() . ']'; }
}
class Formatted {
    use HasFormatter;
    public function __construct(public string $name) {}
    public function format(): string { return $this->name; }
}
echo (new Formatted('ok'))->output();
",
    );
    assert_eq!(out, vec!["[ok]"]);
}

#[test]
fn class_namespace_aware_fqn_runtime() {
    let out = run_prints(
        "<?php
namespace Demo\\Domain;
class Widget {
    public static function id(): string { return __CLASS__; }
}
echo class_alias(Widget::class, __NAMESPACE__ . '\\\\AliasWidget') ? 'alias-ok' : 'alias-fail';
echo '|';
echo AliasWidget::id();
",
    );
    assert_eq!(out, vec!["alias-ok|Demo\\Domain\\AliasWidget"]);
}

#[test]
fn class_dynamic_class_name_instantiation_runtime() {
    let out = run_prints(
        "<?php
class Service {
    public function __construct(public string $name) {}
}
$class = 'Service';
$svc = new $class('core');
echo $svc->name;
",
    );
    assert_eq!(out, vec!["core"]);
}

#[test]
fn class_dynamic_instance_method_call_runtime() {
    let out = run_prints(
        "<?php
class Worker {
    public function ping(): string { return 'pong'; }
}
$obj = new Worker();
$method = 'ping';
echo $obj->$method();
",
    );
    assert_eq!(out, vec!["pong"]);
}

#[test]
fn class_dynamic_static_method_call_runtime() {
    let out = run_prints(
        "<?php
class MathTools {
    public static function double(int $n): int { return $n * 2; }
}
$class = 'MathTools';
$method = 'double';
echo $class::$method(7);
",
    );
    assert_eq!(out, vec!["14"]);
}

#[test]
fn class_call_user_func_array_instance_method_runtime() {
    let out = run_prints(
        "<?php
class Adder {
    public function sum(int $a, int $b): int { return $a + $b; }
}
$obj = new Adder();
$callable = [$obj, 'sum'];
echo call_user_func_array($callable, [4, 6]);
",
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn class_method_chain_with_variable_next_step_runtime() {
    let out = run_prints(
        "<?php
class Pipeline {
    private int $v = 0;
    public function step1(int $n): self { $this->v += $n; return $this; }
    public function step2(string $label): string { return $label . ':' . $this->v; }
}
$p = new Pipeline();
$next = 'step1';
echo $p->$next(3)->step2('ok');
",
    );
    assert_eq!(out, vec!["ok:3"]);
}

#[test]
fn class_dynamic_property_name_runtime() {
    let out = run_prints(
        "<?php
class Holder {
    public int $count = 0;
}
$h = new Holder();
$prop = 'count';
$h->$prop = 11;
echo $h->$prop;
",
    );
    assert_eq!(out, vec!["11"]);
}

#[test]
fn class_magic_call_dynamic_missing_method_runtime() {
    let out = run_prints(
        "<?php
class Dispatcher {
    private array $calls = [];
    public function __call(string $name, array $args): mixed {
        $this->calls[] = $name;
        return $name . '-' . $args[0];
    }
}
$d = new Dispatcher();
echo $d->auto('yes');
",
    );
    assert_eq!(out, vec!["auto-yes"]);
}

#[test]
fn class_dynamic_invokable_object_call_runtime() {
    let out = run_prints(
        "<?php
class Invokable {
    public function __invoke(string $tag): string {
        return strtoupper($tag);
    }
}
$handler = new Invokable();
$fn = $handler;
echo $fn('x');
",
    );
    assert_eq!(out, vec!["X"]);
}

#[test]
fn class_dynamic_call_static_exists_runtime() {
    let out = run_prints(
        "<?php
class Tools {
    public static function make(string $v): string { return 'made:' . $v; }
}
$name = 'make';
echo method_exists(Tools::class, $name) ? 'ok' : 'no';
echo '|';
echo call_user_func([Tools::class, $name], 'cfg');
",
    );
    assert_eq!(out, vec!["ok|made:cfg"]);
}

#[test]
fn class_private_constant_exposed_via_public_accessor_runtime() {
    let out = run_prints(
        "<?php
class Config {
    private const INTERNAL = 'internal';
    public const PUBLIC = 'public';
    public static function internal(): string { return self::INTERNAL; }
}
echo Config::PUBLIC;
echo '|';
echo Config::internal();
",
    );
    assert_eq!(out, vec!["public|internal"]);
}

#[test]
fn class_parent_constructor_and_parent_keyword_runtime() {
    let out = run_prints(
        "<?php
class Base {
    public string $label;
    public function __construct(string $label) { $this->label = $label; }
}
class Child extends Base {
    public function __construct(string $label) {
        parent::__construct('p:' . $label);
    }
}
echo (new Child('x'))->label;
",
    );
    assert_eq!(out, vec!["p:x"]);
}

#[test]
fn class_anonymous_class_extends_named_class_runtime() {
    let out = run_prints(
        "<?php
class Core {
    public function base(): string { return 'base'; }
}
$obj = new class extends Core {
    public function derived(): string { return $this->base() . '-derived'; }
};
echo $obj->derived();
",
    );
    assert_eq!(out, vec!["base-derived"]);
}

#[test]
fn class_implements_iteratoraggregate_runtime() {
    let out = run_prints(
        "<?php
class ListAdapter implements IteratorAggregate {
    private array $items = [1, 2, 3];
    public function getIterator(): Traversable {
        return new ArrayIterator($this->items);
    }
}
$list = new ListAdapter();
echo count(iterator_to_array($list));
",
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn class_static_property_hidden_in_subclass_runtime() {
    let out = run_prints(
        "<?php
class Base {
    public static int $count = 1;
}
class Child extends Base {
    public static int $count = 5;
}
echo Base::$count;
echo '|';
echo Child::$count;
",
    );
    assert_eq!(out, vec!["1|5"]);
}

#[test]
fn class_readonly_property_runtime() {
    let out = run_prints(
        "<?php
class Profile {
    public function __construct(public readonly string $name) {}
}
$p = new Profile('alice');
echo $p->name;
",
    );
    assert_eq!(out, vec!["alice"]);
}

#[test]
fn class_magic_get_and_set_runtime() {
    let out = run_prints(
        "<?php
class Bag {
    private array $values = [];
    public function __set(string $k, mixed $v): void { $this->values[$k] = $v; }
    public function __get(string $k): mixed { return $this->values[$k] ?? null; }
}
$b = new Bag();
$b->item = 'x';
echo $b->item;
",
    );
    assert_eq!(out, vec!["x"]);
}

#[test]
fn class_magic_call_runtime() {
    let out = run_prints(
        "<?php
class MathApi {
    public function __call(string $name, array $args): string {
        return $name . ':' . implode(',', $args);
    }
}
$api = new MathApi();
echo $api->add(2, 3);
",
    );
    assert_eq!(out, vec!["add:2,3"]);
}

#[test]
fn class_invoke_magic_runtime() {
    let out = run_prints(
        "<?php
class Invokable {
    public function __invoke(string $tag): string { return strtoupper($tag); }
}
$inv = new Invokable();
echo $inv('php');
",
    );
    assert_eq!(out, vec!["PHP"]);
}

#[test]
fn class_clone_hook_runtime() {
    let out = run_prints(
        "<?php
class CopyTracker {
    public int $count;
    public function __construct(int $count) { $this->count = $count; }
    public function __clone(): void { $this->count += 1; }
}
$a = new CopyTracker(1);
$b = clone $a;
echo $a->count;
echo '|';
echo $b->count;
",
    );
    assert_eq!(out, vec!["1|2"]);
}

#[test]
fn class_final_class_and_final_method_runtime() {
    let out = run_prints(
        "<?php
class Base {
    final public function id(): string { return 'base'; }
}
class Child extends Base {
}
echo (new Child())->id();
",
    );
    assert_eq!(out, vec!["base"]);
}

#[test]
fn class_implements_countable_iterator_runtime() {
    let out = run_prints(
        "<?php
class Bag implements Countable {
    public function __construct(private array $items) {}
    public function count(): int { return count($this->items); }
}
echo (new Bag([1, 2, 3, 4]))->count();
",
    );
    assert_eq!(out, vec!["4"]);
}

#[test]
fn class_trait_private_method_visibility_runtime() {
    let out = run_prints(
        "<?php
trait Hidden {
    private function label(): string { return 'hidden'; }
}
class Box {
    use Hidden {
        label as public publicLabel;
    }
}
echo (new Box())->publicLabel();
",
    );
    assert_eq!(out, vec!["hidden"]);
}

#[test]
fn class_static_property_visibility_runtime() {
    let out = run_prints(
        "<?php
class Counter {
    protected static int $count = 0;
    public static function inc(): void { self::$count += 1; }
    public static function value(): int { return self::$count; }
}
Counter::inc();
Counter::inc();
echo Counter::value();
",
    );
    assert_eq!(out, vec!["2"]);
}
