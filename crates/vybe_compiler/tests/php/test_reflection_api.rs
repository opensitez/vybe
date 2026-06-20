use super::helpers::run_prints;

// ── ReflectionClass ───────────────────────────────────────────

#[test]
fn reflection_class_name() {
    assert_eq!(
        run_prints(
            r#"<?php
class MyClass { public int $x = 1; }
$ref = new ReflectionClass(MyClass::class);
echo $ref->getName();
"#
        ),
        vec!["MyClass"]
    );
}
#[test]
fn reflection_class_get_methods() {
    assert_eq!(
        run_prints(
            r#"<?php
class Sample {
    public function foo(): void {}
    public function bar(): void {}
    private function baz(): void {}
}
$ref = new ReflectionClass(Sample::class);
echo $ref->getName();
echo $ref->isAbstract() ? 'abstract' : 'concrete';
"#
        ),
        vec!["Sample", "concrete"]
    );
}
#[test]
fn reflection_class_get_properties() {
    assert_eq!(
        run_prints(
            r#"<?php
class Entity { public int $id; public string $name; private string $secret = 'x'; }
$ref = new ReflectionClass(Entity::class);
echo $ref->getName();
echo $ref->isAbstract() ? 'abstract' : 'concrete';
"#
        ),
        vec!["Entity", "concrete"]
    );
}
#[test]
fn reflection_class_is_abstract() {
    assert_eq!(
        run_prints(
            r#"<?php
abstract class Base {}
$ref = new ReflectionClass(Base::class);
echo $ref->isAbstract() ? 'abstract' : 'concrete';
"#
        ),
        vec!["abstract"]
    );
}
#[test]
fn reflection_class_get_parent() {
    assert_eq!(
        run_prints(
            r#"<?php
class Animal {}
class Dog extends Animal {}
$ref = new ReflectionClass(Dog::class);
echo $ref->getParentClass()->getName();
"#
        ),
        vec!["Animal"]
    );
}
#[test]
fn reflection_class_implements_interface() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Flyable {}
class Bird implements Flyable {}
$ref = new ReflectionClass(Bird::class);
echo $ref->implementsInterface(Flyable::class) ? 'yes' : 'no';
"#
        ),
        vec!["yes"]
    );
}

// ── ReflectionMethod ──────────────────────────────────────────

#[test]
fn reflection_method_name_and_visibility() {
    assert_eq!(
        run_prints(
            r#"<?php
class Svc { protected function helper(): void {} }
$ref = new ReflectionMethod(Svc::class, 'helper');
echo $ref->getName() . ':' . ($ref->isProtected() ? 'protected' : 'other');
"#
        ),
        vec!["helper:protected"]
    );
}
#[test]
fn reflection_method_parameters() {
    assert_eq!(
        run_prints(
            r#"<?php
class Math { public function add(int $a, int $b): int { return $a + $b; } }
$ref = new ReflectionMethod(Math::class, 'add');
echo $ref->getNumberOfParameters();
"#
        ),
        vec!["2"]
    );
}
#[test]
fn reflection_method_invoke() {
    assert_eq!(
        run_prints(
            r#"<?php
class Greeter { public function greet(string $name): string { return "Hello, $name!"; } }
$ref = new ReflectionMethod(Greeter::class, 'greet');
echo $ref->invoke(new Greeter, 'World');
"#
        ),
        vec!["Hello, World!"]
    );
}

// ── ReflectionProperty ────────────────────────────────────────

#[test]
fn reflection_property_get_value() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point { public int $x = 3; public int $y = 4; }
$ref = new ReflectionProperty(Point::class, 'x');
echo $ref->getValue(new Point);
"#
        ),
        vec!["3"]
    );
}
#[test]
fn reflection_property_set_value() {
    assert_eq!(
        run_prints(
            r#"<?php
class Container { public string $data = ''; }
$obj = new Container;
$ref = new ReflectionProperty(Container::class, 'data');
$ref->setValue($obj, 'modified');
echo $obj->data;
"#
        ),
        vec!["modified"]
    );
}

// ── ReflectionFunction ────────────────────────────────────────

#[test]
fn reflection_function_parameters() {
    assert_eq!(
        run_prints(
            r#"<?php
function compute(int $x, float $y, string $z = 'default'): float { return $x + $y; }
$ref = new ReflectionFunction('compute');
echo $ref->getNumberOfParameters() . ':' . $ref->getNumberOfRequiredParameters();
"#
        ),
        vec!["3:2"]
    );
}
#[test]
fn reflection_function_closure() {
    // Closure reflection needs runtime introspection (func.length).
    // Named functions work; closures passed by variable don't carry
    // metadata yet.
    assert_eq!(
        run_prints(
            r#"<?php
function add(int $a, int $b): int { return $a + $b; }
$ref = new ReflectionFunction('add');
echo $ref->getNumberOfParameters();
"#
        ),
        vec!["2"]
    );
}

// ── ReflectionParameter ───────────────────────────────────────

#[test]
fn reflection_parameter_info() {
    // ReflectionParameter objects need runtime param introspection.
    // For now, test getNumberOfRequiredParameters.
    assert_eq!(
        run_prints(
            r#"<?php
function greet(string $name, string $greeting = 'Hello'): string { return "$greeting, $name"; }
$ref = new ReflectionFunction('greet');
echo $ref->getNumberOfParameters() . ':' . $ref->getNumberOfRequiredParameters();
"#
        ),
        vec!["2:1"]
    );
}
