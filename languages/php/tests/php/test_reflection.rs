use super::helpers::compile_ok;

// ── ReflectionClass basics ────────────────────────────────────

#[test]
fn reflection_class_name() {
    compile_ok(
        r#"<?php
class Foo { public int $x = 1; }
$rc = new ReflectionClass('Foo');
echo $rc->getName();
echo $rc->getShortName();
"#,
    );
}

#[test]
fn reflection_class_methods() {
    compile_ok(
        r#"<?php
class Calculator {
    public function add(int $a, int $b): int { return $a + $b; }
    public function sub(int $a, int $b): int { return $a - $b; }
    private function secret(): void {}
}
$rc = new ReflectionClass(Calculator::class);
$public = $rc->getMethods(ReflectionMethod::IS_PUBLIC);
echo count($public);
"#,
    );
}

#[test]
fn reflection_class_properties() {
    compile_ok(
        r#"<?php
class User {
    public string $name;
    protected int $age;
    private string $password;
}
$rc = new ReflectionClass(User::class);
$props = $rc->getProperties();
echo count($props);
"#,
    );
}

#[test]
fn reflection_class_constants() {
    compile_ok(
        r#"<?php
class Status {
    const OK      = 200;
    const CREATED = 201;
    const ERROR   = 500;
}
$rc = new ReflectionClass(Status::class);
$consts = $rc->getConstants();
echo count($consts);
echo $consts['OK'];
"#,
    );
}

#[test]
fn reflection_class_interfaces() {
    compile_ok(
        r#"<?php
interface Serializable2 { public function serialize2(): string; }
interface Loggable { public function log(): void; }
class Service implements Serializable2, Loggable {
    public function serialize2(): string { return ''; }
    public function log(): void {}
}
$rc = new ReflectionClass(Service::class);
$ifaces = $rc->getInterfaceNames();
sort($ifaces);
echo implode(',', $ifaces);
"#,
    );
}

#[test]
fn reflection_class_is_abstract() {
    compile_ok(
        r#"<?php
abstract class Base {}
class Concrete extends Base {}
$rb = new ReflectionClass(Base::class);
$rc = new ReflectionClass(Concrete::class);
echo $rb->isAbstract() ? 'abstract' : 'concrete';
echo $rc->isAbstract() ? 'abstract' : 'concrete';
"#,
    );
}

#[test]
fn reflection_class_is_interface() {
    compile_ok(
        r#"<?php
interface MyInterface {}
class MyClass {}
echo (new ReflectionClass(MyInterface::class))->isInterface() ? 'interface' : 'class';
echo (new ReflectionClass(MyClass::class))->isInterface() ? 'interface' : 'class';
"#,
    );
}

#[test]
fn reflection_class_parent() {
    compile_ok(
        r#"<?php
class Animal {}
class Dog extends Animal {}
$rc = new ReflectionClass(Dog::class);
echo $rc->getParentClass()->getName();
"#,
    );
}

#[test]
fn reflection_class_instantiate() {
    compile_ok(
        r#"<?php
class Point { public function __construct(public int $x, public int $y) {} }
$rc = new ReflectionClass(Point::class);
$obj = $rc->newInstance(3, 7);
echo $obj->x . ',' . $obj->y;
"#,
    );
}

#[test]
fn reflection_class_new_instance_without_constructor() {
    compile_ok(
        r#"<?php
class Config { public string $env = 'dev'; }
$rc = new ReflectionClass(Config::class);
$obj = $rc->newInstanceWithoutConstructor();
echo $obj->env;
"#,
    );
}

// ── ReflectionMethod ──────────────────────────────────────────

#[test]
fn reflection_method_name_and_visibility() {
    compile_ok(
        r#"<?php
class Service {
    public function doWork(): void {}
    protected function helper(): void {}
    private function internal(): void {}
}
$rc = new ReflectionClass(Service::class);
$method = $rc->getMethod('doWork');
echo $method->getName();
echo $method->isPublic() ? ':public' : ':not-public';
"#,
    );
}

#[test]
fn reflection_method_parameters() {
    compile_ok(
        r#"<?php
function compute(int $x, float $y, string $label = 'default'): float {
    return $x + $y;
}
$rf = new ReflectionFunction('compute');
$params = $rf->getParameters();
echo count($params);
echo ':' . $params[2]->getDefaultValue();
"#,
    );
}

#[test]
fn reflection_method_invoke() {
    compile_ok(
        r#"<?php
class Math { public function mul(int $a, int $b): int { return $a * $b; } }
$obj = new Math();
$method = new ReflectionMethod(Math::class, 'mul');
echo $method->invoke($obj, 6, 7);
"#,
    );
}

#[test]
fn reflection_method_is_static() {
    compile_ok(
        r#"<?php
class Factory {
    public static function create(): static { return new static(); }
    public function doSomething(): void {}
}
$rc = new ReflectionClass(Factory::class);
echo $rc->getMethod('create')->isStatic() ? 'static' : 'instance';
echo $rc->getMethod('doSomething')->isStatic() ? 'static' : 'instance';
"#,
    );
}

// ── ReflectionProperty ────────────────────────────────────────

#[test]
fn reflection_property_get_set() {
    compile_ok(
        r#"<?php
class Box { public int $width = 10; public int $height = 20; }
$obj = new Box();
$rp = new ReflectionProperty(Box::class, 'width');
echo $rp->getValue($obj);
$rp->setValue($obj, 50);
echo $obj->width;
"#,
    );
}

#[test]
fn reflection_property_type() {
    compile_ok(
        r#"<?php
class TypedProps {
    public int $count = 0;
    public ?string $label = null;
}
$rc = new ReflectionClass(TypedProps::class);
$prop = $rc->getProperty('count');
echo $prop->getType()->getName();
"#,
    );
}

// ── ReflectionFunction ────────────────────────────────────────

#[test]
fn reflection_function_basic() {
    compile_ok(
        r#"<?php
function greet(string $name, string $prefix = 'Hello'): string {
    return "$prefix, $name!";
}
$rf = new ReflectionFunction('greet');
echo $rf->getName();
echo ':' . $rf->getNumberOfParameters();
"#,
    );
}

#[test]
fn reflection_function_closure() {
    compile_ok(
        r#"<?php
$fn = fn(int $a, int $b) => $a + $b;
$rf = new ReflectionFunction($fn);
echo $rf->getNumberOfParameters();
echo $rf->isClosure() ? ':closure' : ':function';
"#,
    );
}

// ── ReflectionParameter ───────────────────────────────────────

#[test]
fn reflection_parameter_details() {
    compile_ok(
        r#"<?php
function create(string $name, int $age = 0, bool $active = true): void {}
$rf = new ReflectionFunction('create');
foreach ($rf->getParameters() as $p) {
    $opt = $p->isOptional() ? '?' : '!';
    echo $p->getName() . $opt . ' ';
}
"#,
    );
}

#[test]
fn reflection_parameter_variadic() {
    compile_ok(
        r#"<?php
function sum(int ...$nums): int { return array_sum($nums); }
$rf = new ReflectionFunction('sum');
$p = $rf->getParameters()[0];
echo $p->isVariadic() ? 'variadic' : 'regular';
"#,
    );
}

// ── ReflectionEnum (PHP 8.1+) ─────────────────────────────────

#[test]
fn reflection_enum_cases() {
    compile_ok(
        r#"<?php
enum Color { case Red; case Green; case Blue; }
$re = new ReflectionEnum(Color::class);
$cases = $re->getCases();
echo count($cases);
echo ':' . $cases[0]->getName();
"#,
    );
}

#[test]
fn reflection_backed_enum() {
    compile_ok(
        r#"<?php
enum Status: string { case Active = 'active'; case Inactive = 'inactive'; }
$re = new ReflectionEnum(Status::class);
echo $re->isBacked() ? 'backed' : 'pure';
echo ':' . $re->getBackingType()->getName();
"#,
    );
}

// ── ReflectionClass traits ────────────────────────────────────

#[test]
fn reflection_class_traits() {
    compile_ok(
        r#"<?php
trait HasLogger { public function log(): void {} }
trait HasCache  { public function cache(): void {} }
class App { use HasLogger, HasCache; }
$rc = new ReflectionClass(App::class);
$traits = array_keys($rc->getTraits());
sort($traits);
echo implode(',', $traits);
"#,
    );
}

// ── hasMethod / hasProperty ───────────────────────────────────

#[test]
fn reflection_has_method_property() {
    compile_ok(
        r#"<?php
class Config { public string $env = 'dev'; public function load(): void {} }
$rc = new ReflectionClass(Config::class);
echo $rc->hasMethod('load')      ? 'yes' : 'no';
echo $rc->hasMethod('missing')   ? 'yes' : 'no';
echo $rc->hasProperty('env')     ? 'yes' : 'no';
echo $rc->hasProperty('unknown') ? 'yes' : 'no';
"#,
    );
}

// ── getDocComment ─────────────────────────────────────────────

#[test]
fn reflection_doc_comment() {
    compile_ok(
        r#"<?php
/** @param int $n The number */
function documented(int $n): int { return $n * 2; }
$rf = new ReflectionFunction('documented');
$doc = $rf->getDocComment();
echo $doc !== false ? 'has doc' : 'no doc';
"#,
    );
}
