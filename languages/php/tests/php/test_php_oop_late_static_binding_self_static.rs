use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Late Static Binding & Static Resolution — static::, self::, parent::, get_called_class(), static factory methods
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_late_static_binding_static_vs_self() {
    let out = run_prints(
        r#"<?php
class BaseService {
    public static function getScopeSelf(): string { return self::getName(); }
    public static function getScopeStatic(): string { return static::getName(); }
    public static function getName(): string { return "Base"; }
}

class ChildService extends BaseService {
    public static function getName(): string { return "Child"; }
}

echo ChildService::getScopeSelf() . " vs " . ChildService::getScopeStatic();
"#,
    );
    assert_eq!(out, vec!["Base vs Child"]);
}

#[test]
fn test_php_static_factory_method_new_static() {
    let out = run_prints(
        r#"<?php
abstract class Model {
    public function __construct(public string $table) {}
    public static function make(): static {
        return new static(static::$defaultTable);
    }
}

class UserModel extends Model {
    protected static string $defaultTable = "users";
}

$user = UserModel::make();
echo get_class($user) . " table={$user->table}";
"#,
    );
    assert_eq!(out, vec!["UserModel table=users"]);
}

#[test]
fn test_php_get_called_class_introspection() {
    let out = run_prints(
        r#"<?php
class ParentClass {
    public static function testClass(): string {
        return get_called_class();
    }
}

class SubClass extends ParentClass {}

echo SubClass::testClass();
"#,
    );
    assert_eq!(out, vec!["SubClass"]);
}

#[test]
fn test_php_static_property_inheritance_and_override() {
    let out = run_prints(
        r#"<?php
class A {
    public static int $count = 10;
}
class B extends A {
    public static int $count = 20;
}

echo A::$count . " vs " . B::$count;
"#,
    );
    assert_eq!(out, vec!["10 vs 20"]);
}

#[test]
fn test_php_forward_static_call_forwarding() {
    compile_ok(
        r#"<?php
class A {
    public static function foo() {
        echo get_called_class();
    }
}
class B extends A {
    public static function foo() {
        forward_static_call(['A', 'foo']);
    }
}

B::foo();
"#,
    );
}

#[test]
fn test_php_static_return_type_hint_php80() {
    compile_ok(
        r#"<?php
class Chainable {
    public function setOption(): static {
        return $this;
    }
}

class SubChainable extends Chainable {}

$sc = new SubChainable();
echo get_class($sc->setOption());
"#,
    );
}

#[test]
fn test_php_static_property_accessed_via_variable() {
    compile_ok(
        r#"<?php
class ConfigHolder {
    public static string $env = "staging";
}

$className = "ConfigHolder";
echo $className::$env;
"#,
    );
}

#[test]
fn test_php_parent_static_method_chaining() {
    compile_ok(
        r#"<?php
class ParentFactory {
    public static function boot() { echo "ParentBoot "; }
}

class ChildFactory extends ParentFactory {
    public static function boot() {
        parent::boot();
        echo "ChildBoot";
    }
}

ChildFactory::boot();
"#,
    );
}

#[test]
fn test_php_static_method_callable_array_syntax() {
    compile_ok(
        r#"<?php
class Dispatcher {
    public static function handle(string $event) { return "Handled: $event"; }
}

$callable = [Dispatcher::class, "handle"];
echo call_user_func($callable, "login");
"#,
    );
}

#[test]
fn test_php_late_static_binding_in_singleton() {
    compile_ok(
        r#"<?php
abstract class Singleton {
    private static array $instances = [];
    public static function getInstance(): static {
        $cls = static::class;
        return self::$instances[$cls] ??= new static();
    }
}

class AppRegistry extends Singleton {}
$reg = AppRegistry::getInstance();
echo get_class($reg);
"#,
    );
}

#[test]
fn test_php_late_static_binding_in_trait_factory() {
    assert_eq!(
        run_prints(
            r#"<?php
trait FactoryTrait {
    public static function make(string $id): static {
        return new static($id);
    }
}

class Service {
    use FactoryTrait;
    public function __construct(public string $id) {}
}

class BillingService extends Service {}

$service = BillingService::make("billing");
echo get_class($service) . "|" . $service->id;
"#,
        ),
        vec!["BillingService|billing"]
    );
}

#[test]
fn test_php_static_binding_parent_static_call_preserves_dynamic_class() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    public static function resolve(): string {
        return static::class;
    }
}

class Mid extends Base {
    public static function callParentResolve(): string {
        return parent::resolve();
    }
}

class Leaf extends Mid {}

echo Leaf::resolve() . "|" . Leaf::callParentResolve();
"#,
        ),
        vec!["Leaf|Leaf"]
    );
}

#[test]
fn test_php_forward_static_calls_static_factory_chain() {
    assert_eq!(
        run_prints(
            r#"<?php
class ParentFactory {
    public static function make(): static {
        return new static();
    }

    public static function label(): string {
        return "parent";
    }
}

class ChildFactory extends ParentFactory {
    public static function label(): string {
        return "child";
    }
}

$factory = ChildFactory::make();
echo get_class($factory) . "|" . $factory::label();
"#,
        ),
        vec!["ChildFactory|child"]
    );
}

#[test]
fn test_php_static_property_shadowing_with_parent_access() {
    assert_eq!(
        run_prints(
            r#"<?php
class BaseCounter {
    protected static int $count = 1;
    public static function current(): int {
        return static::$count;
    }
}

class DerivedCounter extends BaseCounter {
    protected static int $count = 4;
    public static function currentFromParent(): int {
        return parent::$count;
    }
}

echo DerivedCounter::current() . "|" . DerivedCounter::currentFromParent();
"#,
        ),
        vec!["4|1"]
    );
}
