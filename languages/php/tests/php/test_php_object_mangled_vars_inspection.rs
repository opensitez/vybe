use super::helpers::run_prints;

#[test]
fn test_get_mangled_object_vars_public_private_protected() {
    assert_eq!(
        run_prints(
            r#"<?php
class Sample {
    public int $pub = 1;
    protected string $prot = "secret";
    private bool $priv = true;
}

$vars = get_mangled_object_vars(new Sample());
echo count($vars), "\n";
"#
        ),
        vec!["3"]
    );
}

#[test]
fn test_get_mangled_object_vars_mangled_keys() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {
    private int $secret = 42;
}

$vars = get_mangled_object_vars(new Base());
$keys = array_keys($vars);
echo (str_contains($keys[0], 'Base') && str_contains($keys[0], 'secret')) ? 'mangled' : 'plain', "\n";
"#
        ),
        vec!["mangled"]
    );
}

#[test]
fn test_get_mangled_object_vars_dynamic_properties() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new stdClass();
$obj->dynamic = "hello";
$vars = get_mangled_object_vars($obj);
echo $vars['dynamic'], "\n";
"#
        ),
        vec!["hello"]
    );
}

#[test]
fn test_get_mangled_object_vars_protected_prefix() {
    assert_eq!(
        run_prints(
            r#"<?php
class Item {
    protected int $id = 100;
}
$vars = get_mangled_object_vars(new Item());
$keys = array_keys($vars);
echo (str_contains($keys[0], 'id') && str_contains($keys[0], '*')) ? 'protected_prefix' : 'other', "\n";
"#
        ),
        vec!["protected_prefix"]
    );
}

#[test]
fn test_get_mangled_object_vars_uninitialized_typed_property() {
    assert_eq!(
        run_prints(
            r#"<?php
class Typed {
    public string $initialized = "yes";
    public int $uninit;
}
$vars = get_mangled_object_vars(new Typed());
echo array_key_exists('uninit', $vars) ? 'present' : 'absent', "\n";
"#
        ),
        vec!["absent"]
    );
}
