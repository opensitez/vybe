use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Reflection Attribute Instantiation & Targets — ReflectionAttribute::newInstance(), getArguments(), getTarget(), isRepeated()
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_reflection_attribute_arguments_and_instantiation() {
    let out = run_prints(
        r#"<?php
#[Attribute(Attribute::TARGET_CLASS)]
class Table {
    public function __construct(public string $name, public array $indexes = []) {}
}

#[Table("orders", indexes: ["idx_user_id"])]
class Order {}

$rc = new ReflectionClass(Order::class);
$attr = $rc->getAttributes(Table::class)[0];
$args = $attr->getArguments();
$instance = $attr->newInstance();

echo "Name={$instance->name} Index={$instance->indexes[0]} ArgCount=" . count($args);
"#,
    );
    assert_eq!(out, vec!["Name=orders Index=idx_user_id ArgCount=2"]);
}

#[test]
fn test_php_reflection_attribute_target_flags_check() {
    let out = run_prints(
        r#"<?php
#[Attribute(Attribute::TARGET_METHOD | Attribute::TARGET_PROPERTY)]
class Injectable {}

class Controller {
    #[Injectable]
    public string $db;

    #[Injectable]
    public function index() {}
}

$rp = new ReflectionProperty(Controller::class, "db");
$rm = new ReflectionMethod(Controller::class, "index");

echo count($rp->getAttributes(Injectable::class)) . " | " . count($rm->getAttributes(Injectable::class));
"#,
    );
    assert_eq!(out, vec!["1 | 1"]);
}

#[test]
fn test_php_reflection_attribute_repeated_flag() {
    let out = run_prints(
        r#"<?php
#[Attribute(Attribute::TARGET_CLASS | Attribute::IS_REPEATABLE)]
class Tag {
    public function __construct(public string $label) {}
}

#[Tag("auth")]
#[Tag("api")]
#[Tag("v1")]
class ApiService {}

$rc = new ReflectionClass(ApiService::class);
$tags = array_map(fn($a) => $a->newInstance()->label, $rc->getAttributes(Tag::class));
echo implode(", ", $tags);
"#,
    );
    assert_eq!(out, vec!["auth, api, v1"]);
}

#[test]
fn test_php_reflection_attribute_is_instanceof_check() {
    compile_ok(
        r#"<?php
interface BaseAttribute {}

#[Attribute]
class CustomAttr implements BaseAttribute {}

#[CustomAttr]
class Target {}

$rc = new ReflectionClass(Target::class);
$attrs = $rc->getAttributes(BaseAttribute::class, ReflectionAttribute::IS_INSTANCEOF);
echo count($attrs);
"#,
    );
}

#[test]
fn test_php_reflection_attribute_on_function_parameter() {
    compile_ok(
        r#"<?php
#[Attribute(Attribute::TARGET_PARAMETER)]
class ValidateEmail {}

function registerUser(#[ValidateEmail] string $email) {}

$rp = new ReflectionParameter("registerUser", "email");
$attrs = $rp->getAttributes(ValidateEmail::class);
echo count($attrs);
"#,
    );
}

#[test]
fn test_php_reflection_attribute_on_enum_case() {
    compile_ok(
        r#"<?php
#[Attribute(Attribute::TARGET_CLASS_CONSTANT)]
class Label { public function __construct(public string $text) {} }

enum Status {
    #[Label("Pending Approval")]
    case Pending;
}

$re = new ReflectionEnum(Status::class);
$case = $re->getCase("Pending");
$attrs = $case->getAttributes(Label::class);
echo count($attrs);
"#,
    );
}

#[test]
fn test_php_reflection_attribute_on_function() {
    compile_ok(
        r#"<?php
#[Attribute(Attribute::TARGET_FUNCTION)]
class DeprecatedFunction { public function __construct(public string $reason) {} }

#[DeprecatedFunction("Use newApi() instead")]
function oldApi() {}

$rf = new ReflectionFunction("oldApi");
$attr = $rf->getAttributes(DeprecatedFunction::class)[0];
echo $attr->newInstance()->reason;
"#,
    );
}

#[test]
fn test_php_reflection_attribute_name_getter() {
    compile_ok(
        r#"<?php
#[Attribute]
class Component {}

#[Component]
class Service {}

$rc = new ReflectionClass(Service::class);
$attr = $rc->getAttributes()[0];
echo $attr->getName();
"#,
    );
}

#[test]
fn test_php_reflection_attribute_newInstance_error_handling() {
    compile_ok(
        r#"<?php
#[Attribute]
class ConfigAttr { public function __construct(string $required) {} }

#[ConfigAttr] // missing required argument
class Host {}

$rc = new ReflectionClass(Host::class);
$attr = $rc->getAttributes(ConfigAttr::class)[0];
try {
    $attr->newInstance();
} catch (Error $e) {
    echo "Attribute instantiation error caught";
}
"#,
    );
}

#[test]
fn test_php_reflection_attribute_on_anonymous_class() {
    compile_ok(
        r#"<?php
#[Attribute]
class Transient {}

$anon = new #[Transient] class {};
$rc = new ReflectionClass($anon);
echo count($rc->getAttributes(Transient::class));
"#,
    );
}
