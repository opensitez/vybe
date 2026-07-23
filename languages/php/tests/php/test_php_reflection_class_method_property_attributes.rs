use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Reflection API — ReflectionClass, ReflectionMethod, ReflectionProperty, ReflectionAttribute, ReflectionType, ReflectionEnum
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_reflection_class_properties_and_methods() {
    let out = run_prints(
        r#"<?php
class Sample {
    public string $name = "default";
    private int $secret = 42;
    public function execute(): void {}
}

$ref = new ReflectionClass(Sample::class);
echo $ref->getName() . " | " . count($ref->getProperties()) . " | " . count($ref->getMethods());
"#,
    );
    assert_eq!(out, vec!["Sample | 2 | 1"]);
}

#[test]
fn test_php_reflection_method_invoke_and_parameters() {
    let out = run_prints(
        r#"<?php
class Service {
    public function greet(string $name, string $greeting = "Hello"): string {
        return "$greeting $name";
    }
}

$rm = new ReflectionMethod(Service::class, "greet");
$params = $rm->getParameters();
echo $params[0]->getName() . " | " . ($params[1]->isOptional() ? "OPT" : "REQ");
"#,
    );
    assert_eq!(out, vec!["name | OPT"]);
}

#[test]
fn test_php_reflection_property_value_access_by_set_accessible() {
    let out = run_prints(
        r#"<?php
class Account {
    private float $balance = 150.0;
}

$acc = new Account();
$rp = new ReflectionProperty(Account::class, "balance");
echo $rp->getValue($acc);
"#,
    );
    assert_eq!(out, vec!["150"]);
}

#[test]
fn test_php_reflection_attribute_reading() {
    let out = run_prints(
        r#"<?php
#[Attribute]
class Entity {
    public function __construct(public string $table) {}
}

#[Entity("users_table")]
class User {}

$rc = new ReflectionClass(User::class);
$attrs = $rc->getAttributes(Entity::class);
$entity = $attrs[0]->newInstance();
echo $entity->table;
"#,
    );
    assert_eq!(out, vec!["users_table"]);
}

#[test]
fn test_php_reflection_type_union_inspection() {
    let out = run_prints(
        r#"<?php
class Model {
    public int|string $identifier;
}

$rp = new ReflectionProperty(Model::class, "identifier");
$type = $rp->getType();
echo $type->allowsNull() ? "NULLABLE" : "NON_NULL";
"#,
    );
    assert_eq!(out, vec!["NON_NULL"]);
}

#[test]
fn test_php_reflection_enum_cases_inspection() {
    compile_ok(
        r#"<?php
enum Suit: string {
    case Hearts = "H";
    case Diamonds = "D";
}

$re = new ReflectionEnum(Suit::class);
echo $re->isBacked() ? "BACKED" : "PURE";
$cases = $re->getCases();
foreach ($cases as $case) {
    echo $case->getName() . "=" . $case->getValue()->value . "\n";
}
"#,
    );
}

#[test]
fn test_php_reflection_class_doc_comment() {
    compile_ok(
        r#"<?php
/**
 * @Entity(table="products")
 */
class Product {}

$rc = new ReflectionClass(Product::class);
echo str_contains($rc->getDocComment(), "@Entity") ? "DOC_FOUND" : "NO_DOC";
"#,
    );
}

#[test]
fn test_php_reflection_function_invokable() {
    compile_ok(
        r#"<?php
$fn = function(int $a, int $b): int { return $a + $b; };
$rf = new ReflectionFunction($fn);
echo $rf->invoke(10, 20);
"#,
    );
}

#[test]
fn test_php_reflection_class_implements_interface_check() {
    compile_ok(
        r#"<?php
interface PluginInterface {}
class MyPlugin implements PluginInterface {}

$rc = new ReflectionClass(MyPlugin::class);
echo $rc->implementsInterface(PluginInterface::class) ? "IMPLEMENTS" : "NO";
"#,
    );
}

#[test]
fn test_php_reflection_parameter_default_value_available() {
    compile_ok(
        r#"<?php
function testParam($a = 100, $b = "default") {}

$rf = new ReflectionFunction("testParam");
$params = $rf->getParameters();
if ($params[0]->isDefaultValueAvailable()) {
    echo $params[0]->getDefaultValue();
}
"#,
    );
}
