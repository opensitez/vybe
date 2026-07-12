use super::helpers::run_prints;

// ── == vs === for objects ─────────────────────────────────────

#[test]
fn two_separate_objects_not_identical() {
    assert_eq!(
        run_prints(
            r#"<?php
class A {}
$x = new A(); $y = new A();
echo ($x === $y) ? 'identical' : 'not identical';
"#
        ),
        vec!["not identical"]
    );
}

#[test]
fn same_reference_is_identical() {
    assert_eq!(
        run_prints(
            r#"<?php
class A {}
$x = new A(); $y = $x;
echo ($x === $y) ? 'identical' : 'not identical';
"#
        ),
        vec!["identical"]
    );
}

#[test]
fn two_objects_same_class_same_props_are_equal_not_identical() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point { public function __construct(public int $x, public int $y) {} }
$a = new Point(1, 2);
$b = new Point(1, 2);
echo ($a == $b ? 'equal' : 'not equal') . ',' . ($a === $b ? 'identical' : 'not identical');
"#
        ),
        vec!["equal,not identical"]
    );
}

#[test]
fn two_objects_different_props_not_equal() {
    assert_eq!(
        run_prints(
            r#"<?php
class Point { public function __construct(public int $x, public int $y) {} }
$a = new Point(1, 2);
$b = new Point(3, 4);
echo ($a == $b) ? 'equal' : 'not equal';
"#
        ),
        vec!["not equal"]
    );
}

// ── Objects of different classes are never equal ─────────────

#[test]
fn different_class_objects_not_equal_even_with_same_props() {
    assert_eq!(
        run_prints(
            r#"<?php
class Foo { public int $x = 1; }
class Bar { public int $x = 1; }
$a = new Foo(); $b = new Bar();
echo ($a == $b) ? 'equal' : 'not equal';
"#
        ),
        vec!["not equal"]
    );
}

// ── Object == null ────────────────────────────────────────────

#[test]
fn object_not_equal_to_null() {
    assert_eq!(
        run_prints(
            r#"<?php
class A {}
$a = new A();
echo ($a == null) ? 'equal' : 'not equal';
"#
        ),
        vec!["not equal"]
    );
}

#[test]
fn null_not_identical_to_object() {
    assert_eq!(
        run_prints(
            r#"<?php
class A {}
$a = new A();
echo ($a === null) ? 'identical' : 'not identical';
"#
        ),
        vec!["not identical"]
    );
}

// ── clone produces equal but not identical ────────────────────

#[test]
fn cloned_object_equal_but_not_identical() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box { public int $v = 5; }
$a = new Box(); $b = clone $a;
echo ($a == $b ? 'eq' : 'ne') . ',' . ($a === $b ? 'id' : 'nid');
"#
        ),
        vec!["eq,nid"]
    );
}

#[test]
fn modifying_clone_makes_them_unequal() {
    assert_eq!(
        run_prints(
            r#"<?php
class Box { public int $v = 5; }
$a = new Box(); $b = clone $a;
$b->v = 99;
echo ($a == $b) ? 'equal' : 'not equal';
"#
        ),
        vec!["not equal"]
    );
}

// ── Null-safe operator with null ──────────────────────────────

#[test]
fn nullsafe_operator_on_null_returns_null() {
    assert_eq!(
        run_prints(
            r#"<?php
class User { public ?Address $address = null; }
class Address { public string $city = ''; }
$u = new User();
echo var_export($u->address?->city, true);
"#
        ),
        vec!["NULL"]
    );
}

#[test]
fn nullsafe_operator_chain_on_non_null() {
    assert_eq!(
        run_prints(
            r#"<?php
class City { public string $name = 'London'; }
class Address { public City $city; public function __construct() { $this->city = new City(); } }
class User { public Address $address; public function __construct() { $this->address = new Address(); } }
$u = new User();
echo $u->address?->city?->name;
"#
        ),
        vec!["London"]
    );
}

// ── Object comparison in array_unique ────────────────────────

#[test]
fn array_unique_with_equal_objects_keeps_first() {
    assert_eq!(
        run_prints(
            r#"<?php
class Tag { public function __construct(public string $name) {} }
$tags = [new Tag('php'), new Tag('php'), new Tag('rust')];
$unique = array_unique($tags);
echo count($unique);
"#
        ),
        vec!["2"]
    );
}

// ── in_array strict with objects ─────────────────────────────

#[test]
fn in_array_strict_true_for_same_reference() {
    assert_eq!(
        run_prints(
            r#"<?php
class A {}
$obj = new A();
$arr = [$obj, new A()];
echo in_array($obj, $arr, true) ? 'found' : 'not found';
"#
        ),
        vec!["found"]
    );
}

#[test]
fn in_array_strict_false_for_equal_but_different_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class A { public int $v = 1; }
$arr = [new A()];
echo in_array(new A(), $arr, true) ? 'found' : 'not found';
"#
        ),
        vec!["not found"]
    );
}

// ── spl_object_id uniqueness ──────────────────────────────────

#[test]
fn spl_object_id_different_for_different_instances() {
    assert_eq!(
        run_prints(
            r#"<?php
class A {}
$a = new A(); $b = new A();
echo spl_object_id($a) !== spl_object_id($b) ? 'unique' : 'same';
"#
        ),
        vec!["unique"]
    );
}

#[test]
fn spl_object_id_same_for_same_reference() {
    assert_eq!(
        run_prints(
            r#"<?php
class A {}
$a = new A(); $b = $a;
echo spl_object_id($a) === spl_object_id($b) ? 'same' : 'different';
"#
        ),
        vec!["same"]
    );
}

// ── spl_object_hash ───────────────────────────────────────────

#[test]
fn spl_object_hash_same_for_same_reference() {
    assert_eq!(
        run_prints(
            r#"<?php
class A {}
$a = new A(); $b = $a;
echo spl_object_hash($a) === spl_object_hash($b) ? 'same' : 'different';
"#
        ),
        vec!["same"]
    );
}

#[test]
fn spl_object_hash_different_for_clones() {
    assert_eq!(
        run_prints(
            r#"<?php
class A {}
$a = new A(); $b = clone $a;
echo spl_object_hash($a) !== spl_object_hash($b) ? 'different' : 'same';
"#
        ),
        vec!["different"]
    );
}

// ── get_class comparison ──────────────────────────────────────

#[test]
fn get_class_returns_exact_class_name() {
    assert_eq!(
        run_prints(
            r#"<?php
class Foo {}
$f = new Foo();
echo get_class($f);
"#
        ),
        vec!["Foo"]
    );
}

#[test]
fn get_class_returns_child_class_not_parent() {
    assert_eq!(
        run_prints(
            r#"<?php
class Animal {}
class Dog extends Animal {}
$d = new Dog();
echo get_class($d);
"#
        ),
        vec!["Dog"]
    );
}

// ── is_a and instanceof equivalence ──────────────────────────

#[test]
fn is_a_equivalent_to_instanceof_for_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class Animal {}
class Dog extends Animal {}
$d = new Dog();
echo is_a($d, 'Animal') ? 'yes' : 'no';
"#
        ),
        vec!["yes"]
    );
}

#[test]
fn is_a_string_with_allow_string_true() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base {}
class Child extends Base {}
echo is_a('Child', 'Base', true) ? 'yes' : 'no';
"#
        ),
        vec!["yes"]
    );
}

// ── Comparison with == considers property values ──────────────

#[test]
fn object_equality_compares_all_properties() {
    assert_eq!(
        run_prints(
            r#"<?php
class Config { public string $a; public int $b; }
$x = new Config(); $x->a = 'hello'; $x->b = 1;
$y = new Config(); $y->a = 'hello'; $y->b = 1;
$z = new Config(); $z->a = 'hello'; $z->b = 2;
echo ($x == $y ? 'eq' : 'ne') . ',' . ($x == $z ? 'eq' : 'ne');
"#
        ),
        vec!["eq,ne"]
    );
}

// ── Spaceship operator with objects not supported ─────────────

#[test]
fn spaceship_on_equal_objects_returns_zero() {
    assert_eq!(
        run_prints(
            r#"<?php
class Val { public function __construct(public int $n) {} }
$a = new Val(5); $b = new Val(5);
try {
    echo ($a <=> $b);
} catch (\Error $e) {
    echo "error";
}
"#
        ),
        vec!["0"]
    );
}
