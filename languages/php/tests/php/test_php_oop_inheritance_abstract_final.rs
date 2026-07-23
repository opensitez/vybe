use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: OOP Inheritance, Abstract Classes & Final Modifiers — abstract methods, final class/methods, covariant returns, parent::
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_abstract_class_implementation() {
    let out = run_prints(
        r#"<?php
abstract class Shape {
    abstract public function area(): float;
}

class Circle extends Shape {
    public function __construct(public float $radius) {}
    public function area(): float {
        return 3.14159 * $this->radius * $this->radius;
    }
}

$c = new Circle(2.0);
echo round($c->area(), 2);
"#,
    );
    assert_eq!(out, vec!["12.57"]);
}

#[test]
fn test_php_parent_method_invocation() {
    let out = run_prints(
        r#"<?php
class BaseController {
    public function render(): string {
        return "Header";
    }
}

class HomeController extends BaseController {
    public function render(): string {
        return parent::render() . " -> Body";
    }
}

$h = new HomeController();
echo $h->render();
"#,
    );
    assert_eq!(out, vec!["Header -> Body"]);
}

#[test]
fn test_php_final_class_instantiation() {
    let out = run_prints(
        r#"<?php
final class ValueObject {
    public function __construct(public string $value) {}
}

$vo = new ValueObject("immutable");
echo $vo->value;
"#,
    );
    assert_eq!(out, vec!["immutable"]);
}

#[test]
fn test_php_instanceof_type_hierarchy() {
    let out = run_prints(
        r#"<?php
interface Identifiable {}
class Base implements Identifiable {}
class Derived extends Base {}

$d = new Derived();
echo ($d instanceof Derived ? "1" : "0");
echo ($d instanceof Base ? "1" : "0");
echo ($d instanceof Identifiable ? "1" : "0");
"#,
    );
    assert_eq!(out, vec!["111"]);
}

#[test]
fn test_php_covariant_return_type_refinement() {
    compile_ok(
        r#"<?php
class Animal {}
class Dog extends Animal {}

abstract class AnimalFactory {
    abstract public function create(): Animal;
}

class DogFactory extends AnimalFactory {
    public function create(): Dog {
        return new Dog();
    }
}

$df = new DogFactory();
echo get_class($df->create());
"#,
    );
}

#[test]
fn test_php_contravariant_parameter_type_widening() {
    compile_ok(
        r#"<?php
class Cat {}

class Logger {
    public function log(Cat $cat): void {}
}

class UniversalLogger extends Logger {
    public function log(object $entity): void {}
}

$ul = new UniversalLogger();
$ul->log(new stdClass());
"#,
    );
}

#[test]
fn test_php_final_method_declaration() {
    compile_ok(
        r#"<?php
class AuthManager {
    final public function hashPassword(string $pwd): string {
        return md5($pwd);
    }
}

class CustomAuth extends AuthManager {
    public function login(): void {}
}

$ca = new CustomAuth();
echo $ca->hashPassword("12345");
"#,
    );
}

#[test]
fn test_php_abstract_protected_method_override() {
    compile_ok(
        r#"<?php
abstract class DataProcessor {
    abstract protected function transform(array $data): array;
    
    public function process(array $data): array {
        return $this->transform($data);
    }
}

class CSVProcessor extends DataProcessor {
    public function transform(array $data): array {
        return array_map('strtoupper', $data);
    }
}

$p = new CSVProcessor();
print_r($p->process(["a", "b"]));
"#,
    );
}

#[test]
fn test_php_parent_constructor_chaining() {
    compile_ok(
        r#"<?php
class BaseEntity {
    public int $id;
    public function __construct(int $id) {
        $this->id = $id;
    }
}

class UserEntity extends BaseEntity {
    public string $email;
    public function __construct(int $id, string $email) {
        parent::__construct($id);
        $this->email = $email;
    }
}

$u = new UserEntity(1, "user@example.com");
echo "{$u->id}: {$u->email}";
"#,
    );
}

#[test]
fn test_php_interface_multiple_inheritance() {
    compile_ok(
        r#"<?php
interface Readable { public function read(): string; }
interface Writable { public function write(string $data): void; }
interface ReadWriteable extends Readable, Writable {}

class Buffer implements ReadWriteable {
    private string $content = "";
    public function read(): string { return $this->content; }
    public function write(string $data): void { $this->content .= $data; }
}

$b = new Buffer();
$b->write("hello");
echo $b->read();
"#,
    );
}
