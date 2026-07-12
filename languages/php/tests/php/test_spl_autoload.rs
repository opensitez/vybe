//! `spl_autoload_register` and class/interface/trait existence checks.

crate::php_cases! {
    spl_autoload_register_loads_namespaced_class => {
        r#"<?php
spl_autoload_register(function (string $class): void {
    if ($class === 'App\\Widget') {
        eval('namespace App; class Widget { public function id(): string { return "w1"; } }');
    }
});
echo (new App\Widget())->id();
"#,
        ["w1"]
    };

    spl_autoload_unregister_leaves_class_loaded => {
        r#"<?php
$loader = function (string $class): void {
    if ($class === 'Tmp\\Once') {
        eval('namespace Tmp; class Once {}');
    }
};
spl_autoload_register($loader);
class_exists('Tmp\\Once');
spl_autoload_unregister($loader);
echo class_exists('Tmp\\Once', false) ? 'loaded' : 'gone';
"#,
        ["loaded"]
    };

    class_exists_without_autoload_returns_false_for_missing => {
        r#"<?php
echo class_exists('Missing\\Class', false) ? 'yes' : 'no';
"#,
        ["no"]
    };

    class_exists_after_definition_returns_true => {
        r#"<?php
class LocalSvc {}
echo class_exists('LocalSvc', false) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    interface_exists_detects_declared_interface => {
        r#"<?php
interface Capable { public function run(): void; }
echo interface_exists('Capable', false) ? 'iface' : 'no';
"#,
        ["iface"]
    };

    trait_exists_detects_declared_trait => {
        r#"<?php
trait Loggable { public function log(): string { return 'log'; } }
echo trait_exists('Loggable', false) ? 'trait' : 'no';
"#,
        ["trait"]
    };

    method_exists_on_instance => {
        r#"<?php
class Svc { public function handle(): void {} }
echo method_exists(new Svc(), 'handle') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    method_exists_missing_returns_false => {
        r#"<?php
class Svc {}
echo method_exists('Svc', 'missing') ? 'yes' : 'no';
"#,
        ["no"]
    };

    property_exists_checks_declared_property => {
        r#"<?php
class Box { public string $label = 'x'; }
echo property_exists(new Box(), 'label') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    get_declared_classes_includes_stdclass => {
        r#"<?php
echo in_array('stdClass', get_declared_classes(), true) ? 'std' : 'no';
"#,
        ["std"]
    };

    get_declared_interfaces_includes_stringable => {
        r#"<?php
echo interface_exists('Stringable', false) ? 'str' : 'no';
"#,
        ["str"]
    };

    get_declared_traits_lists_user_trait => {
        r#"<?php
trait Marker {}
echo in_array('Marker', get_declared_traits(), true) ? 'marked' : 'no';
"#,
        ["marked"]
    };

    autoload_stack_invokes_registered_loader => {
        r#"<?php
$hit = false;
spl_autoload_register(function (string $class) use (&$hit): void {
    if ($class === 'Hit\\Load') { $hit = true; }
});
class_exists('Hit\\Load');
echo $hit ? 'hit' : 'miss';
"#,
        ["hit"]
    };

    class_alias_creates_alias_name => {
        r#"<?php
class RealThing {}
class_alias(RealThing::class, 'AliasThing');
echo (new AliasThing()) instanceof RealThing ? 'alias' : 'no';
"#,
        ["alias"]
    };

    get_parent_class_reports_extends => {
        r#"<?php
class ParentCls {}
class ChildCls extends ParentCls {}
echo get_parent_class(ChildCls::class);
"#,
        ["ParentCls"]
    };

    is_subclass_of_detects_inheritance => {
        r#"<?php
class Base {}
class Derived extends Base {}
echo is_subclass_of(Derived::class, Base::class) ? 'sub' : 'no';
"#,
        ["sub"]
    };

    is_a_with_string_class_name => {
        r#"<?php
class Node {}
$n = new Node();
echo is_a($n, 'Node') ? 'node' : 'other';
"#,
        ["node"]
    };
}
