use super::helpers::{compile_ok, run_prints};

// ── Simple variable interpolation edge cases ──────────────────

#[test]
fn interpolation_variable_adjacent_text() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = "world";
echo "hello{$x}end";
"#
        ),
        vec!["helloworldend"]
    );
}

#[test]
fn interpolation_integer_variable() {
    assert_eq!(
        run_prints(
            r#"<?php
$n = 42;
echo "value is $n units";
"#
        ),
        vec!["value is 42 units"]
    );
}

#[test]
fn interpolation_boolean_variable() {
    assert_eq!(
        run_prints(
            r#"<?php
$b = true;
echo "flag: $b";
"#
        ),
        vec!["flag: 1"]
    );
}

// ── Array access interpolation ────────────────────────────────

#[test]
fn interpolation_simple_array_index() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = ['a', 'b', 'c'];
echo "item: $arr[1]";
"#
        ),
        vec!["item: b"]
    );
}

#[test]
fn interpolation_array_string_key_curly() {
    assert_eq!(
        run_prints(
            r#"<?php
$map = ['name' => 'Alice'];
echo "user: {$map['name']}";
"#
        ),
        vec!["user: Alice"]
    );
}

#[test]
fn interpolation_nested_array_curly() {
    assert_eq!(
        run_prints(
            r#"<?php
$data = [['x' => 99]];
echo "val: {$data[0]['x']}";
"#
        ),
        vec!["val: 99"]
    );
}

#[test]
fn interpolation_array_in_expression_not_interpolated_without_curly() {
    compile_ok(
        r#"<?php
$arr = [1, 2, 3];
$s = "count is " . count($arr);
echo $s;
"#,
    );
}

// ── Object property interpolation ─────────────────────────────

#[test]
fn interpolation_object_property_curly() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new stdClass();
$obj->color = "red";
echo "color: {$obj->color}";
"#
        ),
        vec!["color: red"]
    );
}

#[test]
fn interpolation_object_chained_property() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = new stdClass();
$a->b = new stdClass();
$a->b->val = "deep";
echo "val: {$a->b->val}";
"#
        ),
        vec!["val: deep"]
    );
}

#[test]
fn interpolation_object_array_property() {
    assert_eq!(
        run_prints(
            r#"<?php
$obj = new stdClass();
$obj->items = ['first', 'second'];
echo "item: {$obj->items[0]}";
"#
        ),
        vec!["item: first"]
    );
}

// ── Variable variable interpolation ──────────────────────────

#[test]
fn interpolation_variable_variable_curly() {
    assert_eq!(
        run_prints(
            r#"<?php
$varname = "greeting";
$$varname = "hello";
echo "say: ${$varname}";
"#
        ),
        vec!["say: hello"]
    );
}

// ── Method call NOT interpolated directly ─────────────────────

#[test]
fn interpolation_method_call_requires_concat() {
    assert_eq!(
        run_prints(
            r#"<?php
class Greeter { public function greet(): string { return "hi"; } }
$g = new Greeter();
echo "result: " . $g->greet();
"#
        ),
        vec!["result: hi"]
    );
}

// ── Curly-dollar vs dollar-curly distinction ──────────────────

#[test]
fn interpolation_dollar_curly_vs_curly_dollar() {
    assert_eq!(
        run_prints(
            r#"<?php
$fruit = "apple";
$apple = "red apple";
echo "${fruit}";
echo " ";
echo "{$fruit}";
"#
        ),
        vec!["apple red apple"]
    );
}

// ── Escape sequences inside double-quoted strings ─────────────

#[test]
fn interpolation_newline_escape() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = "line1\nline2";
echo substr_count($s, "\n");
"#
        ),
        vec!["1"]
    );
}

#[test]
fn interpolation_tab_escape() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = "a\tb";
echo strlen($s);
"#
        ),
        vec!["3"]
    );
}

#[test]
fn interpolation_unicode_escape_sequence() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = "\u{0041}";
echo $s;
"#
        ),
        vec!["A"]
    );
}

#[test]
fn interpolation_hex_escape_sequence() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = "\x41\x42\x43";
echo $s;
"#
        ),
        vec!["ABC"]
    );
}

#[test]
fn interpolation_octal_escape_sequence() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = "\101";
echo $s;
"#
        ),
        vec!["A"]
    );
}

#[test]
fn interpolation_backslash_dollar_not_interpolated() {
    assert_eq!(
        run_prints(
            r#"<?php
$price = 5;
echo "cost: \$price";
"#
        ),
        vec!["cost: $price"]
    );
}

// ── Complex expression sequences ──────────────────────────────

#[test]
fn interpolation_multiple_variables_same_string() {
    assert_eq!(
        run_prints(
            r#"<?php
$first = "John";
$last = "Doe";
echo "$first $last";
"#
        ),
        vec!["John Doe"]
    );
}

#[test]
fn interpolation_arithmetic_outside_string() {
    assert_eq!(
        run_prints(
            r#"<?php
$n = 5;
echo "double: " . ($n * 2);
"#
        ),
        vec!["double: 10"]
    );
}

#[test]
fn interpolation_ternary_result_concatenated() {
    assert_eq!(
        run_prints(
            r#"<?php
$score = 80;
echo "grade: " . ($score >= 60 ? "pass" : "fail");
"#
        ),
        vec!["grade: pass"]
    );
}

#[test]
fn interpolation_null_variable_becomes_empty() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = null;
echo "val:$x:end";
"#
        ),
        vec!["val::end"]
    );
}

#[test]
fn interpolation_float_variable() {
    assert_eq!(
        run_prints(
            r#"<?php
$pi = 3.14;
echo "pi=$pi";
"#
        ),
        vec!["pi=3.14"]
    );
}

#[test]
fn interpolation_string_inside_string_is_literal() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = "inner";
echo "outer '$s' end";
"#
        ),
        vec!["outer 'inner' end"]
    );
}

#[test]
fn interpolation_array_index_zero() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['zero', 'one'];
echo "first: $a[0]";
"#
        ),
        vec!["first: zero"]
    );
}
