use super::helpers::run_prints;

// ── (int) cast ────────────────────────────────────────────────

#[test]
fn cast_float_to_int_truncates() {
    assert_eq!(run_prints(r#"<?php echo (int)3.9; "#), vec!["3"]);
}
#[test]
fn cast_negative_float_to_int_truncates_toward_zero() {
    assert_eq!(run_prints(r#"<?php echo (int)-3.9; "#), vec!["-3"]);
}
#[test]
fn cast_numeric_string_to_int() {
    assert_eq!(run_prints(r#"<?php echo (int)'42abc'; "#), vec!["42"]);
}
#[test]
fn cast_bool_true_to_int() {
    assert_eq!(
        run_prints(r#"<?php echo (int)true . ',' . (int)false; "#),
        vec!["1,0"]
    );
}
#[test]
fn cast_null_to_int() {
    assert_eq!(run_prints(r#"<?php echo (int)null; "#), vec!["0"]);
}

// ── (float) cast ──────────────────────────────────────────────

#[test]
fn cast_string_to_float() {
    assert_eq!(run_prints(r#"<?php echo (float)'3.14'; "#), vec!["3.14"]);
}
#[test]
fn cast_int_to_float() {
    assert_eq!(
        run_prints(r#"<?php $f = (float)5; echo is_float($f) ? 'float' : 'int'; "#),
        vec!["float"]
    );
}
#[test]
fn cast_nonnumeric_string_to_float_zero() {
    assert_eq!(run_prints(r#"<?php echo (float)'hello'; "#), vec!["0"]);
}

// ── (string) cast ─────────────────────────────────────────────

#[test]
fn cast_int_to_string() {
    assert_eq!(
        run_prints(r#"<?php $s = (string)42; echo gettype($s) . ':' . $s; "#),
        vec!["string:42"]
    );
}
#[test]
fn cast_bool_to_string() {
    assert_eq!(
        run_prints(r#"<?php echo (string)true . '|' . (string)false; "#),
        vec!["1|"]
    );
}
#[test]
fn cast_null_to_string() {
    assert_eq!(
        run_prints(r#"<?php $s = (string)null; echo strlen($s); "#),
        vec!["0"]
    );
}
#[test]
fn cast_float_to_string_format() {
    assert_eq!(run_prints(r#"<?php echo (string)1.5; "#), vec!["1.5"]);
}

// ── (bool) cast ───────────────────────────────────────────────

#[test]
fn cast_zero_to_bool_false() {
    assert_eq!(
        run_prints(r#"<?php echo (bool)0 ? 'true' : 'false'; "#),
        vec!["false"]
    );
}
#[test]
fn cast_empty_string_to_bool_false() {
    assert_eq!(
        run_prints(r#"<?php echo (bool)'' ? 'true' : 'false'; "#),
        vec!["false"]
    );
}
#[test]
fn cast_string_zero_to_bool_false() {
    assert_eq!(
        run_prints(r#"<?php echo (bool)'0' ? 'true' : 'false'; "#),
        vec!["false"]
    );
}
#[test]
fn cast_nonempty_string_to_bool_true() {
    assert_eq!(
        run_prints(r#"<?php echo (bool)'false' ? 'true' : 'false'; "#),
        vec!["true"]
    );
}
#[test]
fn cast_empty_array_to_bool_false() {
    assert_eq!(
        run_prints(r#"<?php echo (bool)[] ? 'true' : 'false'; "#),
        vec!["false"]
    );
}

// ── (array) cast ──────────────────────────────────────────────

#[test]
fn cast_scalar_to_array_wraps() {
    assert_eq!(
        run_prints(r#"<?php $a = (array)42; echo $a[0]; "#),
        vec!["42"]
    );
}
#[test]
fn cast_null_to_array_empty() {
    assert_eq!(
        run_prints(r#"<?php $a = (array)null; echo count($a); "#),
        vec!["0"]
    );
}
#[test]
fn cast_object_to_array_gets_props() {
    assert_eq!(
        run_prints(
            r#"<?php
class Pt { public int $x = 3; public int $y = 4; }
$a = (array)(new Pt);
echo $a['x'] . ',' . $a['y'];
"#
        ),
        vec!["3,4"]
    );
}

// ── (object) cast ─────────────────────────────────────────────

#[test]
fn cast_array_to_object() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = ['name' => 'Alice', 'age' => 30];
$obj = (object)$arr;
echo $obj->name . ':' . $obj->age;
"#
        ),
        vec!["Alice:30"]
    );
}
#[test]
fn cast_object_to_object_same_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class Foo { public int $v = 7; }
$f = new Foo;
$o = (object)$f;
echo $o->v;
"#
        ),
        vec!["7"]
    );
}

// ── settype() and gettype() ───────────────────────────────────

#[test]
fn settype_converts_in_place() {
    assert_eq!(
        run_prints(r#"<?php $v = '42'; settype($v, 'integer'); echo gettype($v) . ':' . $v; "#),
        vec!["integer:42"]
    );
}
#[test]
fn gettype_values() {
    assert_eq!(
        run_prints(
            r#"<?php
echo gettype(42) . ',' . gettype(3.14) . ',' . gettype('x') . ',' . gettype(true) . ',' . gettype(null) . ',' . gettype([]);
"#
        ),
        vec!["integer,double,string,boolean,NULL,array"]
    );
}
