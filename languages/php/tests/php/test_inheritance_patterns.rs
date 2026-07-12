use super::helpers::run_prints;

// ── Constructor inheritance ───────────────────────────────────

#[test]
fn parent_constructor_called() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base { protected string $type; public function __construct(string $t) { $this->type = $t; } }
class Child extends Base { public function __construct() { parent::__construct('child'); } }
$c = new Child;
echo $c->type, "\n";
"#
        ),
        vec!["child"]
    );
}
#[test]
fn constructor_chain_three_levels() {
    assert_eq!(
        run_prints(
            r#"<?php
class A { public string $name = ''; public function __construct(string $n) { $this->name = $n; } }
class B extends A { public int $id = 0; public function __construct(int $id) { parent::__construct("B-$id"); $this->id = $id; } }
class C extends B { public function __construct() { parent::__construct(42); } }
$c = new C;
echo $c->name . ':' . $c->id, "\n";
"#
        ),
        vec!["B-42:42"]
    );
}

// ── Method overriding ─────────────────────────────────────────

#[test]
fn override_method_calls_parent() {
    assert_eq!(
        run_prints(
            r#"<?php
class Logger {
    public function log(string $msg): string { return "[$msg]"; }
}
class TimestampLogger extends Logger {
    public function log(string $msg): string { return parent::log("2024:" . $msg); }
}
echo (new TimestampLogger)->log('test'), "\n";
"#
        ),
        vec!["[2024:test]"]
    );
}
#[test]
fn override_in_deep_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
class A { public function describe(): string { return 'A'; } }
class B extends A { public function describe(): string { return parent::describe() . 'B'; } }
class C extends B { public function describe(): string { return parent::describe() . 'C'; } }
echo (new C)->describe(), "\n";
"#
        ),
        vec!["ABC"]
    );
}

// ── Polymorphism ──────────────────────────────────────────────

#[test]
fn polymorphic_dispatch() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Payment { abstract public function pay(float $amount): string; }
class CreditCard extends Payment { public function pay(float $a): string { return "CC:$a"; } }
class PayPal extends Payment { public function pay(float $a): string { return "PP:$a"; } }
$payments = [new CreditCard, new PayPal, new CreditCard];
echo implode(',', array_map(fn($p) => $p->pay(10.0), $payments)), "\n";
"#
        ),
        vec!["CC:10,PP:10,CC:10"]
    );
}
#[test]
fn instanceof_in_hierarchy() {
    assert_eq!(
        run_prints(
            r#"<?php
class Vehicle {}
class Car extends Vehicle {}
class ElectricCar extends Car {}
$e = new ElectricCar;
echo ($e instanceof ElectricCar ? '1' : '0') . ($e instanceof Car ? '1' : '0') . ($e instanceof Vehicle ? '1' : '0'), "\n";
"#
        ),
        vec!["111"]
    );
}

// ── final prevents override ───────────────────────────────────

#[test]
fn final_method_cannot_override() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base { final public function sealed(): string { return 'sealed'; } }
class Child extends Base {}
echo (new Child)->sealed(), "\n";
"#
        ),
        vec!["sealed"]
    );
}
#[test]
fn final_class_cannot_extend() {
    assert_eq!(
        run_prints(
            r#"<?php
final class FinalClass {}
try { eval('class Attempt extends FinalClass {}'); } catch (Error $e) { echo 'blocked', "\n"; }
"#
        ),
        vec!["blocked"]
    );
}

// ── Abstract method implementation ────────────────────────────

#[test]
fn abstract_method_in_concrete_subclass() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Serializer {
    abstract protected function encode(mixed $data): string;
    public function serialize(mixed $data): string { return $this->encode($data); }
}
class JsonSerializer extends Serializer {
    protected function encode(mixed $data): string { return json_encode($data); }
}
echo (new JsonSerializer)->serialize(['key' => 'val']), "\n";
"#
        ),
        vec!["{\"key\":\"val\"}"]
    );
}

// ── Interface + abstract class combo ─────────────────────────

#[test]
fn interface_and_abstract_class() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Cacheable { public function cacheKey(): string; }
abstract class BaseModel implements Cacheable {
    abstract public function id(): int;
    public function cacheKey(): string { return get_class($this) . ':' . $this->id(); }
}
class User4 extends BaseModel { public function id(): int { return 42; } }
echo (new User4)->cacheKey(), "\n";
"#
        ),
        vec!["User4:42"]
    );
}

// ── get_class / is_a / instanceof ────────────────────────────

#[test]
fn get_class_hierarchy() {
    assert_eq!(
        run_prints(
            r#"<?php
class A {} class B extends A {} class C extends B {}
$c = new C;
echo get_class($c) . ':' . get_parent_class($c) . ':' . is_a($c, 'A') ? 'a' : 'not', "\n";
"#
        ),
        vec!["a"]
    );
}
#[test]
fn is_subclass_of() {
    assert_eq!(
        run_prints(
            r#"<?php
class Animal {} class Mammal extends Animal {} class Dog extends Mammal {}
echo is_subclass_of(Dog::class, Animal::class) ? 'yes' : 'no', "\n";
echo is_subclass_of(Animal::class, Dog::class) ? 'yes' : 'no', "\n";
"#
        ),
        // Each echo ends with "\n"; the harness splits stdout into lines.
        vec!["yes", "no"]
    );
}

// ── Magic method in inheritance ───────────────────────────────

#[test]
#[allow(non_snake_case)]
fn toString_inherited() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base2 { public function __toString(): string { return 'Base'; } }
class Child2 extends Base2 {}
echo new Child2, "\n";
"#
        ),
        vec!["Base"]
    );
}
#[test]
#[allow(non_snake_case)]
fn toString_overridden() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base3 { public function __toString(): string { return 'Base'; } }
class Child3 extends Base3 { public function __toString(): string { return 'Child'; } }
echo new Child3, "\n";
"#
        ),
        vec!["Child"]
    );
}

// ── Property visibility in inheritance ───────────────────────

#[test]
fn protected_property_accessed_in_child() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base4 { protected int $val = 10; }
class Child4 extends Base4 { public function getVal(): int { return $this->val; } }
echo (new Child4)->getVal(), "\n";
"#
        ),
        vec!["10"]
    );
}
#[test]
fn private_not_visible_in_child() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base5 { private int $secret = 42; }
class Child5 extends Base5 { public function try(): string { return isset($this->secret) ? 'visible' : 'hidden'; } }
echo (new Child5)->try(), "\n";
"#
        ),
        vec!["hidden"]
    );
}
