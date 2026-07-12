//! Property read/write/isset on `null`, scalars, and non-objects.

crate::php_cases! {
    read_property_on_null_throws_error => {
        r#"<?php
$x = null;
try { echo $x->name; }
catch (Error $e) { echo 'null-read'; }
"#,
        ["null-read"]
    };

    write_property_on_null_throws_error => {
        r#"<?php
$x = null;
try { $x->name = 1; echo 'ok'; }
catch (Error $e) { echo 'null-write'; }
"#,
        ["null-write"]
    };

    isset_property_on_null_is_false_without_fatal => {
        r#"<?php
$x = null;
echo isset($x->name) ? 'yes' : 'no';
"#,
        ["no"]
    };

    empty_property_on_null_is_true => {
        r#"<?php
$x = null;
echo empty($x->name) ? 'empty' : 'set';
"#,
        ["empty"]
    };

    read_property_on_false_throws_type_error => {
        r#"<?php
$x = false;
try { echo $x->field; }
catch (TypeError $e) { echo 'false-read'; }
"#,
        ["false-read"]
    };

    read_property_on_int_throws_type_error => {
        r#"<?php
try { echo 1->n; }
catch (TypeError $e) { echo 'int-read'; }
"#,
        ["int-read"]
    };

    read_property_on_array_throws_type_error => {
        r#"<?php
$a = [1];
try { echo $a->x; }
catch (TypeError $e) { echo 'arr-read'; }
"#,
        ["arr-read"]
    };

    write_property_on_array_throws_type_error => {
        r#"<?php
$a = [];
try { $a->x = 1; echo 'ok'; }
catch (TypeError $e) { echo 'arr-write'; }
"#,
        ["arr-write"]
    };

    method_call_on_null_throws_error => {
        r#"<?php
$x = null;
try { $x->run(); echo 'ok'; }
catch (Error $e) { echo 'null-call'; }
"#,
        ["null-call"]
    };

    static_call_on_non_class_string_throws_error => {
        r#"<?php
try { 'not-a-class'::go(); echo 'ok'; }
catch (Error $e) { echo 'bad-static'; }
"#,
        ["bad-static"]
    };

    clone_on_non_object_throws_type_error => {
        r#"<?php
try { clone 1; echo 'ok'; }
catch (TypeError $e) { echo 'clone-int'; }
"#,
        ["clone-int"]
    };

    instanceof_with_non_object_left_hand_side => {
        r#"<?php
echo (1 instanceof stdClass) ? 'yes' : 'no';
"#,
        ["no"]
    };

    property_exists_on_null_returns_false => {
        r#"<?php
echo property_exists(null, 'x') ? 'yes' : 'no';
"#,
        ["no"]
    };

    get_object_vars_on_non_object_throws => {
        r#"<?php
try { get_object_vars(1); echo 'ok'; }
catch (TypeError $e) { echo 'gov-int'; }
"#,
        ["gov-int"]
    };

    get_class_on_non_object_without_false_flag_throws => {
        r#"<?php
try { get_class(1); echo 'ok'; }
catch (TypeError $e) { echo 'gc-int'; }
"#,
        ["gc-int"]
    };

    get_class_on_non_object_with_false_returns_false => {
        r#"<?php
echo get_class(1, false) === false ? 'false' : 'name';
"#,
        ["false"]
    };

    unset_property_on_null_throws_error => {
        r#"<?php
$x = null;
try { unset($x->p); echo 'ok'; }
catch (Error $e) { echo 'null-unset'; }
"#,
        ["ok"]
    };

    indirect_call_on_null_callable_throws => {
        r#"<?php
$f = null;
try { $f(); echo 'ok'; }
catch (TypeError $e) { echo 'null-invoke'; }
"#,
        ["null-invoke"]
    };

    read_private_property_from_outside_triggers_error => {
        r#"<?php
class Vault { private string $secret = 'hidden'; }
$v = new Vault();
try { echo $v->secret; }
catch (Error $e) { echo 'private'; }
"#,
        ["private"]
    };

    write_dynamic_property_on_false_throws => {
        r#"<?php
$x = false;
try { $x->dyn = 1; echo 'ok'; }
catch (TypeError $e) { echo 'false-write'; }
"#,
        ["false-write"]
    };
}
