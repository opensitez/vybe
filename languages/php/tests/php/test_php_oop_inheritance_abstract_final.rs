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

#[test]
fn test_php_abstract_method_dispatch_runtime() {
    let out = run_prints(
        r#"<?php
abstract class Processor {
    abstract protected function transform(string $input): string;
    public function run(string $input): string {
        return $this->transform($input) . '!';
    }
}

class UpperProcessor extends Processor {
    protected function transform(string $input): string {
        return strtoupper($input);
    }
}

echo (new UpperProcessor())->run('ok');
"#,
    );
    assert_eq!(out, vec!["OK!"]);
}

#[test]
fn test_php_static_binding_factory_runtime() {
    let out = run_prints(
        r#"<?php
abstract class Base {
    public function __construct(public string $name) {}
    public static function make(string $name): static {
        return new static($name);
    }
}
class ChildBase extends Base {}
echo ChildBase::make('x')->name;
"#,
    );
    assert_eq!(out, vec!["x"]);
}

#[test]
fn test_php_trait_precedence_runtime() {
    let out = run_prints(
        r#"<?php
trait A {
    public function label(): string { return 'A'; }
}
trait B {
    public function label(): string { return 'B'; }
}
class Item {
    use A, B { B::label insteadof A; }
}
echo (new Item())->label();
"#,
    );
    assert_eq!(out, vec!["B"]);
}

#[test]
fn test_php_final_class_as_dependency_runtime() {
    compile_ok(
        r#"<?php
final class FinalService {
    public function execute(): string { return 'ok'; }
}
class Wrapped {
    public function call(FinalService $s): string { return $s->execute(); }
}
$w = new Wrapped();
echo $w->call(new FinalService());
"#,
    );
}

#[test]
fn test_php_interface_default_method_style_polymorphism() {
    let out = run_prints(
        r#"<?php
interface Sink {
    public function label(): string;
}
class A implements Sink {
    public function label(): string { return 'A'; }
}
class B implements Sink {
    public function label(): string { return 'B'; }
}
$xs = [new A(), new B()];
echo $xs[0]->label() . $xs[1]->label();
"#,
    );
    assert_eq!(out, vec!["AB"]);
}

#[test]
fn test_php_self_static_class_binding_runtime() {
    let out = run_prints(
        r#"<?php
abstract class BaseType {
    public static function selfClass(): string { return self::class; }
    public static function staticClass(): string { return static::class; }
}
class ChildType extends BaseType {}
echo BaseType::selfClass();
echo '|';
echo BaseType::staticClass();
echo '|';
echo ChildType::selfClass();
echo '|';
echo ChildType::staticClass();
"#,
    );
    assert_eq!(out, vec!["BaseType|BaseType|BaseType|ChildType"]);
}

#[test]
fn test_php_abstract_template_method_runtime() {
    let out = run_prints(
        r#"<?php
abstract class Workflow {
    abstract protected function stepA(): string;
    abstract protected function stepB(): string;
    final public function run(): string {
        return $this->stepA() . '>' . $this->stepB();
    }
}
class Job extends Workflow {
    protected function stepA(): string { return 'prepare'; }
    protected function stepB(): string { return 'execute'; }
}
echo (new Job())->run();
"#,
    );
    assert_eq!(out, vec!["prepare>execute"]);
}

#[test]
fn test_php_abstract_property_visibility_preserved_runtime() {
    compile_ok(
        r#"<?php
abstract class UserEntity {
    public function __construct(protected string $name) {}
    protected function getName(): string { return $this->name; }
}
class CustomerEntity extends UserEntity {
    public function label(): string { return $this->getName(); }
}
echo (new CustomerEntity('acme'))->label();
"#,
    );
}

#[test]
fn test_php_trait_alias_runtime() {
    let out = run_prints(
        r#"<?php
trait First {
    public function kind(): string { return 'first'; }
}
trait Second {
    public function kind(): string { return 'second'; }
}
class Item {
    use First, Second {
        First::kind insteadof Second;
        Second::kind as secondaryKind;
    }
}
$item = new Item();
echo $item->kind();
echo '|';
echo $item->secondaryKind();
"#,
    );
    assert_eq!(out, vec!["first|second"]);
}

#[test]
fn test_php_parent_and_static_in_chain_runtime() {
    let out = run_prints(
        r#"<?php
class A {
    public function who(): string { return 'A'; }
}
class B extends A {
    public function who(): string { return 'B' . parent::who(); }
}
class C extends B {
    public function who(): string { return 'C' . parent::who(); }
}
echo (new C())->who();
"#,
    );
    assert_eq!(out, vec!["CBA"]);
}

#[test]
fn test_php_final_class_dependency_runtime() {
    let out = run_prints(
        r#"<?php
interface Repository { public function id(): string; }
final class SqlRepository implements Repository {
    public function __construct(private string $name) {}
    public function id(): string { return $this->name; }
}
function describe(Repository $repo): string { return $repo->id(); }
echo describe(new SqlRepository('main'));
"#,
    );
    assert_eq!(out, vec!["main"]);
}
