use super::helpers::compile_ok;

// ── interface_exists ──────────────────────────────────────────────
#[test]
fn interface_exists_defined_interface() {
    compile_ok(
        r#"<?php
interface Drawable { public function draw(): void; }
echo interface_exists('Drawable') ? 'yes' : 'no';
echo interface_exists('Nonexistent') ? 'yes' : 'no';
"#,
    );
}

// ── trait_exists ──────────────────────────────────────────────────
#[test]
fn trait_exists_defined_trait() {
    compile_ok(
        r#"<?php
trait Greetable { public function greet() { return 'hi'; } }
echo trait_exists('Greetable') ? 'yes' : 'no';
echo trait_exists('Missing') ? 'yes' : 'no';
"#,
    );
}

// ── get_class_methods ────────────────────────────────────────────
#[test]
fn get_class_methods_returns_array() {
    compile_ok(
        r#"<?php
class Calc {
    public function add($a, $b) { return $a + $b; }
    public function sub($a, $b) { return $a - $b; }
}
$methods = get_class_methods('Calc');
echo in_array('add', $methods) ? 'yes' : 'no';
echo in_array('sub', $methods) ? 'yes' : 'no';
"#,
    );
}

// ── get_class_vars ───────────────────────────────────────────────
#[test]
fn get_class_vars_default_properties() {
    compile_ok(
        r#"<?php
class Config {
    public string $host = 'localhost';
    public int $port = 8080;
}
$vars = get_class_vars('Config');
echo isset($vars['host']) ? 'yes' : 'no';
echo isset($vars['port']) ? 'yes' : 'no';
"#,
    );
}

// ── get_object_vars ──────────────────────────────────────────────
#[test]
fn get_object_vars_public_properties() {
    compile_ok(
        r#"<?php
class Point {
    public int $x;
    public int $y;
    public function __construct(int $x, int $y) {
        $this->x = $x;
        $this->y = $y;
    }
}
$p = new Point(3, 7);
$vars = get_object_vars($p);
echo $vars['x'];
echo $vars['y'];
"#,
    );
}

// ── get_parent_class ─────────────────────────────────────────────
#[test]
fn get_parent_class_with_inheritance() {
    compile_ok(
        r#"<?php
class Base {}
class Child extends Base {}
echo get_parent_class(new Child());
echo get_parent_class('Child');
"#,
    );
}

// ── get_called_class ─────────────────────────────────────────────
#[test]
fn get_called_class_static_context() {
    compile_ok(
        r#"<?php
class ParentClass {
    public static function whoAmI(): string {
        return get_called_class();
    }
}
class ChildClass extends ParentClass {}
echo ChildClass::whoAmI();
"#,
    );
}

// ── is_subclass_of with string class name ────────────────────────
#[test]
fn is_subclass_of_string_class_arg() {
    compile_ok(
        r#"<?php
class Shape {}
class Circle extends Shape {}
echo is_subclass_of('Circle', 'Shape') ? 'yes' : 'no';
echo is_subclass_of('Shape', 'Circle') ? 'yes' : 'no';
"#,
    );
}

// ── class_implements ─────────────────────────────────────────────
#[test]
fn class_implements_interfaces_list() {
    compile_ok(
        r#"<?php
interface Serializable { public function serialize(): string; }
interface Loggable { public function log(): void; }
class Entity implements Serializable, Loggable {
    public function serialize(): string { return ''; }
    public function log(): void {}
}
$ifaces = class_implements('Entity');
echo isset($ifaces['Serializable']) ? 'yes' : 'no';
echo isset($ifaces['Loggable']) ? 'yes' : 'no';
"#,
    );
}

// ── class_uses ───────────────────────────────────────────────────
#[test]
fn class_uses_trait_list() {
    compile_ok(
        r#"<?php
trait HasUuid { public function uuid(): string { return 'abc'; } }
trait HasTimestamp { public function ts(): int { return 0; } }
class Record {
    use HasUuid, HasTimestamp;
}
$traits = class_uses('Record');
echo isset($traits['HasUuid']) ? 'yes' : 'no';
echo isset($traits['HasTimestamp']) ? 'yes' : 'no';
"#,
    );
}

// ── class_parents ────────────────────────────────────────────────
#[test]
fn class_parents_hierarchy() {
    compile_ok(
        r#"<?php
class A {}
class B extends A {}
class C extends B {}
$parents = class_parents('C');
echo isset($parents['B']) ? 'yes' : 'no';
echo isset($parents['A']) ? 'yes' : 'no';
"#,
    );
}

// ── property_exists ──────────────────────────────────────────────
#[test]
fn property_exists_class_and_object() {
    compile_ok(
        r#"<?php
class User {
    public string $name;
    protected int $age = 0;
}
$u = new User();
$u->name = 'Alice';
echo property_exists($u, 'name') ? 'yes' : 'no';
echo property_exists($u, 'age') ? 'yes' : 'no';
echo property_exists($u, 'email') ? 'yes' : 'no';
"#,
    );
}

// ── method_exists ────────────────────────────────────────────────
#[test]
fn method_exists_class_and_object() {
    compile_ok(
        r#"<?php
class Greeter {
    public function hello(): string { return 'hi'; }
}
$g = new Greeter();
echo method_exists($g, 'hello') ? 'yes' : 'no';
echo method_exists($g, 'goodbye') ? 'yes' : 'no';
echo method_exists('Greeter', 'hello') ? 'yes' : 'no';
"#,
    );
}

// ── class_exists with autoload=false ────────────────────────────
#[test]
fn class_exists_no_autoload() {
    compile_ok(
        r#"<?php
class RegisteredClass {}
echo class_exists('RegisteredClass', false) ? 'yes' : 'no';
echo class_exists('UnregisteredClass', false) ? 'yes' : 'no';
"#,
    );
}

// ── get_class with no argument ───────────────────────────────────
#[test]
fn get_class_no_arg_from_method() {
    compile_ok(
        r#"<?php
class SelfNaming {
    public function className(): string {
        return get_class($this);
    }
}
$obj = new SelfNaming();
echo $obj->className();
"#,
    );
}
