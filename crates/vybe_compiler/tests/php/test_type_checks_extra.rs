use super::helpers::compile_ok;

// ── is_scalar ────────────────────────────────────────────────────
#[test]
fn is_scalar_int_value() { compile_ok(r#"<?php
echo is_scalar(42) ? 'yes' : 'no';
echo is_scalar([1, 2]) ? 'yes' : 'no';
"#); }

// ── is_countable ─────────────────────────────────────────────────
#[test]
fn is_countable_array_and_string() { compile_ok(r#"<?php
echo is_countable([1, 2, 3]) ? 'yes' : 'no';
echo is_countable('hello') ? 'yes' : 'no';
"#); }

// ── is_iterable ──────────────────────────────────────────────────
#[test]
fn is_iterable_array_check() { compile_ok(r#"<?php
echo is_iterable([1, 2, 3]) ? 'yes' : 'no';
echo is_iterable(42) ? 'yes' : 'no';
"#); }

// ── is_resource ──────────────────────────────────────────────────
#[test]
fn is_resource_non_resource_value() { compile_ok(r#"<?php
$x = 42;
echo is_resource($x) ? 'yes' : 'no';
"#); }

// ── var_export output ────────────────────────────────────────────
#[test]
fn var_export_simple_values() { compile_ok(r#"<?php
var_export(42);
var_export(3.14);
var_export(true);
var_export(null);
"#); }

// ── var_export with return=true ──────────────────────────────────
#[test]
fn var_export_return_true_assign() { compile_ok(r#"<?php
$s = var_export([1, 2, 3], true);
echo $s;
"#); }

// ── get_defined_constants ────────────────────────────────────────
#[test]
fn get_defined_constants_has_entries() { compile_ok(r#"<?php
define('MY_CONST', 99);
$consts = get_defined_constants(true);
echo isset($consts['user']['MY_CONST']) ? 'yes' : 'no';
"#); }

// ── get_defined_functions ────────────────────────────────────────
#[test]
fn get_defined_functions_user_key() { compile_ok(r#"<?php
function myFunc() { return 1; }
$fns = get_defined_functions();
echo isset($fns['user']) ? 'yes' : 'no';
"#); }

// ── debug_backtrace ──────────────────────────────────────────────
#[test]
fn debug_backtrace_compile_ok() { compile_ok(r#"<?php
function inner() {
    $bt = debug_backtrace();
    return count($bt);
}
function outer() {
    return inner();
}
echo outer();
"#); }

// ── debug_print_backtrace ────────────────────────────────────────
#[test]
fn debug_print_backtrace_compile_ok() { compile_ok(r#"<?php
function traced() {
    debug_print_backtrace();
}
traced();
"#); }

// ── print_r with return=true ─────────────────────────────────────
#[test]
fn print_r_return_true_assign() { compile_ok(r#"<?php
$arr = ['a' => 1, 'b' => 2];
$out = print_r($arr, true);
echo strlen($out) > 0 ? 'has output' : 'empty';
"#); }

// ── var_dump multiple values ─────────────────────────────────────
#[test]
fn var_dump_multiple_values() { compile_ok(r#"<?php
var_dump(42, 'hello', true, null, 3.14);
"#); }

// ── is_subclass_of ───────────────────────────────────────────────
#[test]
fn is_subclass_of_inheritance_check() { compile_ok(r#"<?php
class Animal {}
class Dog extends Animal {}
$d = new Dog();
echo is_subclass_of($d, 'Animal') ? 'yes' : 'no';
echo is_subclass_of($d, 'Dog') ? 'yes' : 'no';
"#); }

// ── is_a ─────────────────────────────────────────────────────────
#[test]
fn is_a_object_and_string_class() { compile_ok(r#"<?php
class Cat {}
class Kitten extends Cat {}
$k = new Kitten();
echo is_a($k, 'Cat') ? 'yes' : 'no';
echo is_a($k, 'Kitten') ? 'yes' : 'no';
echo is_a('Kitten', 'Cat', true) ? 'yes' : 'no';
"#); }

// ── get_defined_vars in function scope ───────────────────────────
#[test]
fn get_defined_vars_local_scope() { compile_ok(r#"<?php
function checkVars() {
    $x = 10;
    $y = 20;
    $vars = get_defined_vars();
    echo isset($vars['x']) ? 'yes' : 'no';
    echo isset($vars['y']) ? 'yes' : 'no';
}
checkVars();
"#); }
