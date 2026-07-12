//! `get_class`, `method_exists`, `property_exists`, and class introspection helpers.

crate::php_cases! {
    get_class_returns_class_name => {
        r#"<?php
class Worker {}
echo get_class(new Worker());
"#,
        ["Worker"]
    };

    get_class_on_object_lowercase_false => {
        r#"<?php
class Api {}
echo get_class(new Api()) === 'Api' ? 'match' : 'diff';
"#,
        ["match"]
    };

    get_parent_class_returns_base => {
        r#"<?php
class Base {}
class Child extends Base {}
echo get_parent_class('Child');
"#,
        ["Base"]
    };

    is_subclass_of_true_for_child => {
        r#"<?php
class P {}
class C extends P {}
echo is_subclass_of('C', 'P') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_a_instance_check => {
        r#"<?php
class T {}
echo is_a(new T(), T::class) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    method_exists_instance_method => {
        r#"<?php
class Svc { public function run(): void {} }
echo method_exists('Svc', 'run') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    method_exists_missing_method => {
        r#"<?php
class Svc {}
echo method_exists('Svc', 'missing') ? 'yes' : 'no';
"#,
        ["no"]
    };

    property_exists_public_property => {
        r#"<?php
class Box { public int $n = 1; }
echo property_exists(new Box(), 'n') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    property_exists_dynamic_after_set => {
        r#"<?php
$o = new stdClass();
$o->dyn = 1;
echo property_exists($o, 'dyn') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    class_exists_after_definition => {
        r#"<?php
class TmpCls {}
echo class_exists('TmpCls') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    interface_exists_after_definition => {
        r#"<?php
interface Cap {}
echo interface_exists('Cap') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    trait_exists_after_definition => {
        r#"<?php
trait T {}
echo trait_exists('T') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    get_declared_classes_includes_stdclass => {
        r#"<?php
echo in_array('stdClass', get_declared_classes(), true) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    get_object_vars_public_only => {
        r#"<?php
class V { public int $a = 1; private int $b = 2; }
echo json_encode(get_object_vars(new V()));
"#,
        ["{\"a\":1}"]
    };

    get_class_methods_lists_public_method => {
        r#"<?php
class M { public function go(): void {} }
echo in_array('go', get_class_methods('M'), true) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    instanceof_operator_true => {
        r#"<?php
class Node {}
echo (new Node()) instanceof Node ? 'yes' : 'no';
"#,
        ["yes"]
    };

    is_object_on_instance => {
        r#"<?php
echo is_object(new stdClass()) ? 'yes' : 'no';
"#,
        ["yes"]
    };

    enum_exists_for_declared_enum => {
        r#"<?php
enum E { case A; }
echo enum_exists('E') ? 'yes' : 'no';
"#,
        ["yes"]
    };
}
