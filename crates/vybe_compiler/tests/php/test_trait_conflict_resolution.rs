use super::helpers::{compile_ok, run_prints};

// ── insteadof — pick one trait's method over another ──────────

#[test] fn insteadof_resolves_method_conflict() {
    assert_eq!(run_prints(r#"<?php
trait A { public function hello(): string { return "A"; } }
trait B { public function hello(): string { return "B"; } }
class C {
    use A, B { A::hello insteadof B; }
}
echo (new C())->hello();
"#), vec!["A"]);
}

#[test] fn insteadof_picks_second_trait() {
    assert_eq!(run_prints(r#"<?php
trait A { public function greet(): string { return "from A"; } }
trait B { public function greet(): string { return "from B"; } }
class C {
    use A, B { B::greet insteadof A; }
}
echo (new C())->greet();
"#), vec!["from B"]);
}

#[test] fn insteadof_multiple_conflicts_resolved() {
    assert_eq!(run_prints(r#"<?php
trait X { public function foo(): string { return "X:foo"; } public function bar(): string { return "X:bar"; } }
trait Y { public function foo(): string { return "Y:foo"; } public function bar(): string { return "Y:bar"; } }
class Z {
    use X, Y {
        X::foo insteadof Y;
        Y::bar insteadof X;
    }
}
$z = new Z();
echo $z->foo() . ',' . $z->bar();
"#), vec!["X:foo,Y:bar"]);
}

// ── as — alias a conflict-excluded method ─────────────────────

#[test] fn as_alias_allows_accessing_excluded_method() {
    assert_eq!(run_prints(r#"<?php
trait A { public function hello(): string { return "A"; } }
trait B { public function hello(): string { return "B"; } }
class C {
    use A, B {
        A::hello insteadof B;
        B::hello as helloFromB;
    }
}
$c = new C();
echo $c->hello() . ',' . $c->helloFromB();
"#), vec!["A,B"]);
}

#[test] fn as_alias_without_conflict() {
    assert_eq!(run_prints(r#"<?php
trait Logger { public function log(string $msg): void { echo $msg; } }
class App {
    use Logger { log as writeLog; }
}
(new App())->writeLog("hello");
"#), vec!["hello"]);
}

// ── as — visibility change ────────────────────────────────────

#[test] fn as_changes_method_visibility_to_protected() {
    compile_ok(r#"<?php
trait Impl { public function secret(): string { return "s"; } }
class Service {
    use Impl { secret as protected; }
}
"#);
}

#[test] fn as_changes_method_visibility_to_private() {
    compile_ok(r#"<?php
trait Impl { public function helper(): string { return "h"; } }
class Widget {
    use Impl { helper as private internalHelper; }
    public function run(): string { return $this->internalHelper(); }
}
echo (new Widget())->run();
"#);
}

// ── Abstract method in trait ──────────────────────────────────

#[test] fn trait_abstract_method_must_be_implemented() {
    assert_eq!(run_prints(r#"<?php
trait Validatable {
    abstract protected function rules(): array;
    public function validate(array $data): bool {
        foreach ($this->rules() as $field) {
            if (empty($data[$field])) return false;
        }
        return true;
    }
}
class Form {
    use Validatable;
    protected function rules(): array { return ['name', 'email']; }
}
$f = new Form();
echo $f->validate(['name' => 'Alice', 'email' => 'a@b.com']) ? 'valid' : 'invalid';
"#), vec!["valid"]);
}

#[test] fn trait_abstract_method_fails_validation_on_missing_field() {
    assert_eq!(run_prints(r#"<?php
trait Validatable {
    abstract protected function required(): array;
    public function check(array $data): string {
        foreach ($this->required() as $field) {
            if (!isset($data[$field])) return "missing: $field";
        }
        return "ok";
    }
}
class UserForm {
    use Validatable;
    protected function required(): array { return ['name', 'age']; }
}
echo (new UserForm())->check(['name' => 'Bob']);
"#), vec!["missing: age"]);
}

// ── Trait with properties ─────────────────────────────────────

#[test] fn trait_property_accessible_in_class() {
    assert_eq!(run_prints(r#"<?php
trait HasName { public string $name = ''; }
class Person { use HasName; }
$p = new Person();
$p->name = "Carol";
echo $p->name;
"#), vec!["Carol"]);
}

#[test] fn trait_static_property_shared_in_class() {
    assert_eq!(run_prints(r#"<?php
trait Counter { public static int $count = 0; }
class Widget { use Counter; }
Widget::$count++;
Widget::$count++;
echo Widget::$count;
"#), vec!["2"]);
}

// ── Multiple traits without conflict ─────────────────────────

#[test] fn multiple_traits_no_conflict_compose() {
    assert_eq!(run_prints(r#"<?php
trait Loggable { public function log(): string { return "log"; } }
trait Cacheable { public function cache(): string { return "cache"; } }
class Service { use Loggable, Cacheable; }
$s = new Service();
echo $s->log() . ',' . $s->cache();
"#), vec!["log,cache"]);
}

// ── Trait in abstract class ───────────────────────────────────

#[test] fn trait_used_in_abstract_class() {
    assert_eq!(run_prints(r#"<?php
trait Greeter { public function greet(): string { return "Hello, " . $this->getName(); } }
abstract class Base { use Greeter; abstract public function getName(): string; }
class Concrete extends Base { public function getName(): string { return "World"; } }
echo (new Concrete())->greet();
"#), vec!["Hello, World"]);
}

// ── Trait calling $this methods ───────────────────────────────

#[test] fn trait_method_calls_class_method() {
    assert_eq!(run_prints(r#"<?php
trait Formatter {
    public function format(): string { return "[" . $this->raw() . "]"; }
}
class Item {
    use Formatter;
    public function raw(): string { return "data"; }
}
echo (new Item())->format();
"#), vec!["[data]"]);
}

// ── Trait with constants (PHP 8.2) ────────────────────────────

#[test] fn trait_with_constant_php82() {
    compile_ok(r#"<?php
trait HasVersion {
    const VERSION = '1.0';
}
class App {
    use HasVersion;
}
echo App::VERSION;
"#);
}

// ── Trait inheritance chain ───────────────────────────────────

#[test] fn class_using_trait_can_still_extend() {
    assert_eq!(run_prints(r#"<?php
trait Taggable { public function tag(): string { return "tagged"; } }
class Base { public function base(): string { return "base"; } }
class Child extends Base { use Taggable; }
$c = new Child();
echo $c->base() . ',' . $c->tag();
"#), vec!["base,tagged"]);
}

// ── Trait method overriding by class ─────────────────────────

#[test] fn class_method_overrides_trait_method() {
    assert_eq!(run_prints(r#"<?php
trait DefaultGreet { public function greet(): string { return "trait"; } }
class Custom { use DefaultGreet; public function greet(): string { return "class"; } }
echo (new Custom())->greet();
"#), vec!["class"]);
}

// ── Trait used in interface implementation ────────────────────

#[test] fn trait_satisfies_interface_requirement() {
    assert_eq!(run_prints(r#"<?php
interface Printable { public function print(): void; }
trait PrintImpl { public function print(): void { echo "printed"; } }
class Doc implements Printable { use PrintImpl; }
(new Doc())->print();
"#), vec!["printed"]);
}

// ── Multiple trait aliases ────────────────────────────────────

#[test] fn multiple_as_aliases_on_same_method() {
    assert_eq!(run_prints(r#"<?php
trait Source { public function value(): int { return 42; } }
class Consumer {
    use Source {
        value as getValue;
        value as fetchValue;
    }
}
$c = new Consumer();
echo $c->getValue() . ',' . $c->fetchValue();
"#), vec!["42,42"]);
}

// ── Trait static method ───────────────────────────────────────

#[test] fn trait_static_method_callable_on_class() {
    assert_eq!(run_prints(r#"<?php
trait Factory {
    public static function create(): static { return new static(); }
}
class Product { use Factory; public function name(): string { return "product"; } }
echo Product::create()->name();
"#), vec!["product"]);
}

// ── Trait with constructor-style initialization ───────────────

#[test] fn trait_init_method_called_from_constructor() {
    assert_eq!(run_prints(r#"<?php
trait Initializable {
    private bool $initialized = false;
    protected function init(): void { $this->initialized = true; }
    public function isReady(): bool { return $this->initialized; }
}
class Service {
    use Initializable;
    public function __construct() { $this->init(); }
}
echo (new Service())->isReady() ? 'ready' : 'not ready';
"#), vec!["ready"]);
}

// ── Trait property conflict detection ────────────────────────

#[test] fn trait_same_property_compatible_redeclaration() {
    compile_ok(r#"<?php
trait HasId { public int $id = 0; }
class Entity {
    use HasId;
    public int $id = 0;
}
"#);
}
