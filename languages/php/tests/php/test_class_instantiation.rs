//! `new` on types that cannot be constructed — abstract, interface, trait, enum, missing classes.

crate::php_cases! {
    constructor_can_instantiate_class_declared_later => {
        r#"<?php
class Project {
    public function __construct() { $this->workflow = new Workflow(); }
    public function label(): string { return $this->workflow->label(); }
}
class Workflow {
    public function label(): string { return 'W'; }
}
$p = new Project();
echo $p->label();
"#,
        ["W"]
    };

    instantiate_abstract_class_throws_error => {
        r#"<?php
abstract class Abs {}
try { new Abs(); echo 'ok'; }
catch (Error $e) { echo 'abstract'; }
"#,
        ["abstract"]
    };

    instantiate_interface_throws_error => {
        r#"<?php
interface Port {}
try { new Port(); echo 'ok'; }
catch (Error $e) { echo 'iface'; }
"#,
        ["iface"]
    };

    instantiate_trait_throws_error => {
        r#"<?php
trait Mix {}
try { new Mix(); echo 'ok'; }
catch (Error $e) { echo 'trait'; }
"#,
        ["trait"]
    };

    instantiate_enum_type_throws_error => {
        r#"<?php
enum Color { case Red; }
try { new Color(); echo 'ok'; }
catch (Error $e) { echo 'enum'; }
"#,
        ["enum"]
    };

    instantiate_enum_case_directly_throws_error => {
        r#"<?php
enum Color { case Red; }
try { new Color::Red; echo 'ok'; }
catch (Error $e) { echo 'case'; }
"#,
        ["case"]
    };

    instantiate_undefined_class_throws_error => {
        r#"<?php
try { new MissingClass(); echo 'ok'; }
catch (Error $e) { echo 'missing'; }
"#,
        ["missing"]
    };

    private_constructor_blocks_external_new => {
        r#"<?php
class Vault {
    private function __construct() {}
    public static function open(): self { return new self(); }
}
try { new Vault(); echo 'ok'; }
catch (Error $e) { echo 'private'; }
"#,
        ["private"]
    };

    static_factory_can_call_private_constructor => {
        r#"<?php
class Vault {
    private function __construct() {}
    public static function open(): self { return new self(); }
}
echo Vault::open() instanceof Vault ? 'factory' : 'no';
"#,
        ["factory"]
    };

    protected_constructor_blocks_global_new => {
        r#"<?php
class Base {
    protected function __construct() {}
}
try { new Base(); echo 'ok'; }
catch (Error $e) { echo 'protected'; }
"#,
        ["protected"]
    };

    child_can_call_protected_parent_constructor => {
        r#"<?php
class Base { protected function __construct(public string $tag) {} }
class Child extends Base { public function __construct() { parent::__construct('kid'); } }
echo (new Child())->tag;
"#,
        ["kid"]
    };

    instantiate_final_class_via_child_is_allowed => {
        r#"<?php
final class Leaf {}
echo (new Leaf()) instanceof Leaf ? 'leaf' : 'no';
"#,
        ["leaf"]
    };

    extend_final_class_at_parse_time_fails => {
        r#"<?php
try {
    eval('final class F {} class G extends F {}');
    echo 'ok';
} catch (Error $e) {
    echo 'final';
}
"#,
        ["final"]
    };

    cannot_new_abstract_method_holder_without_concrete_child => {
        r#"<?php
abstract class Worker { abstract public function run(): string; }
try { new Worker(); echo 'ok'; }
catch (Error $e) { echo 'abs-worker'; }
"#,
        ["abs-worker"]
    };

    concrete_child_of_abstract_is_allowed => {
        r#"<?php
abstract class Worker { abstract public function run(): string; }
class Job extends Worker { public function run(): string { return 'done'; } }
echo (new Job())->run();
"#,
        ["done"]
    };

    instantiate_class_name_from_variable => {
        r#"<?php
$class = 'stdClass';
echo (new $class()) instanceof stdClass ? 'dynamic' : 'no';
"#,
        ["dynamic"]
    };

    instantiate_undefined_class_from_variable => {
        r#"<?php
$name = 'NoSuch';
try { new $name(); echo 'ok'; }
catch (Error $e) { echo 'dyn-miss'; }
"#,
        ["dyn-miss"]
    };

    clone_requires_object_not_scalar => {
        r#"<?php
try { clone 1; echo 'ok'; }
catch (TypeError $e) { echo 'clone-scalar'; }
"#,
        ["clone-scalar"]
    };

    clone_on_uncloneable_internal_may_fail => {
        r#"<?php
$fp = fopen('php://memory', 'r+');
try { clone $fp; echo 'ok'; }
catch (Error $e) { echo 'clone-res'; }
finally { fclose($fp); }
"#,
        ["clone-res"]
    };

    instantiate_self_from_static_scope => {
        r#"<?php
class Seed {
    public function __construct(public string $name) {}
    public static function make(string $name): self {
        return new self($name);
    }
}
echo Seed::make('root')->name;
"#,
        ["root"]
    };

    instantiate_via_variable_with_qualifier => {
        r#"<?php
namespace Demo;
class Worker { public string $role = 'ok'; }
$class_name = __NAMESPACE__ . '\\\\' . 'Worker';
$obj = new $class_name;
echo $obj->role;
"#,
        ["ok"]
    };

    instantiate_with_parenthesized_new_class_name => {
        r#"<?php
class Boxy {
    public function __construct(public string $v) {}
}
$parts = ['B', 'o', 'x', 'y'];
$name = implode('', $parts);
$obj = new $name('x');
echo $obj->v;
"#,
        ["x"]
    };

    instantiate_from_eval_defined_class => {
        r#"<?php
eval('class RuntimeBuilt {}');
echo (new RuntimeBuilt()) instanceof RuntimeBuilt ? 'run' : 'no';
"#,
        ["run"]
    };

    instantiate_trait_string_as_class_must_fail => {
        r#"<?php
trait TraitLike {}
$name = 'TraitLike';
try { new $name(); echo 'ok'; }
catch (Error $e) { echo 'trait-var'; }
"#,
        ["trait-var"]
    };
}
