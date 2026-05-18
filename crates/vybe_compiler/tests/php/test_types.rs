use super::helpers::{compile_ok, run_prints};

// ── Type coercion / juggling ────────────────────────────────
#[test] fn string_to_int_add() { compile_ok("<?php $x = '5' + 3; echo $x;"); }
#[test] fn string_to_int_sub() { compile_ok("<?php $x = '10' - 3; echo $x;"); }
#[test] fn string_to_float() { compile_ok("<?php $x = '3.14' + 0; echo $x;"); }
#[test] fn bool_to_int() { compile_ok("<?php $x = true + true + false; echo $x;"); }
#[test] fn null_to_int() { compile_ok("<?php $x = null + 5; echo $x;"); }
#[test] fn concat_coerce() { compile_ok("<?php $x = 'count: ' . 42; echo $x;"); }
#[test] fn comparison_coerce() { compile_ok("<?php $x = '0' == false; $y = '' == false; $z = '1' == true;"); }

// ── Type checking ───────────────────────────────────────────
#[test] fn is_null() { compile_ok("<?php echo is_null(null); echo is_null(0); echo is_null('');"); }
#[test] fn is_numeric() { compile_ok("<?php echo is_numeric(42); echo is_numeric('3.14'); echo is_numeric('abc');"); }
#[test] fn is_string() { compile_ok("<?php echo is_string('hi'); echo is_string(42);"); }
#[test] fn is_int() { compile_ok("<?php echo is_int(42); echo is_int(3.14);"); }
#[test] fn is_bool() { compile_ok("<?php echo is_bool(true); echo is_bool(0);"); }
#[test] fn is_array() { compile_ok("<?php echo is_array([]); echo is_array('x');"); }
#[test] fn gettype_check() { compile_ok("<?php echo gettype(42); echo gettype('hi'); echo gettype(null); echo gettype(true); echo gettype([]);"); }

// ── Casting ─────────────────────────────────────────────────
#[test] fn cast_int() { compile_ok("<?php $x = (int)'42'; echo $x;"); }
#[test] fn cast_integer_long_form() { compile_ok("<?php $x = (integer)trim(' 42 '); echo $x;"); }
#[test] fn cast_float() { compile_ok("<?php $x = (float)'3.14'; echo $x;"); }
#[test] fn cast_string() { compile_ok("<?php $x = (string)42; echo $x;"); }
#[test] fn cast_bool() { compile_ok("<?php $x = (bool)''; $y = (bool)'hello';"); }
#[test] fn cast_boolean_long_form_function_call() { compile_ok(r#"<?php
class Reader { public $currentTagContents = " 1 "; }
$reader = new Reader();
$value = (boolean)trim($reader->currentTagContents);
echo $value;
"#); }
#[test] fn cast_array() { compile_ok("<?php $x = (array)'hello';"); }

// ── intval / floatval / strval / boolval ────────────────────
#[test] fn intval_string() { compile_ok("<?php echo intval('42abc');"); }
#[test] fn floatval_string() { compile_ok("<?php echo floatval('3.14xyz');"); }
#[test] fn strval_number() { compile_ok("<?php echo strval(42);"); }
#[test] fn boolval_values() { compile_ok("<?php echo boolval(0); echo boolval(1); echo boolval(''); echo boolval('x');"); }

// ── isset / empty / unset ───────────────────────────────────
#[test] fn isset_defined() { compile_ok("<?php $x = 1; echo isset($x);"); }
#[test] fn isset_null() { compile_ok("<?php $x = null; echo isset($x);"); }
#[test] fn isset_multi() { compile_ok("<?php $a = 1; $b = 2; echo isset($a, $b);"); }
#[test] fn empty_values() {
	assert_eq!(run_prints("<?php echo empty('') ? 't' : 'f'; echo empty(0) ? 't' : 'f'; echo empty(null) ? 't' : 'f'; echo empty('0') ? 't' : 'f'; echo empty('x') ? 't' : 'f'; echo empty([]) ? 't' : 'f';"), &["t", "t", "t", "t", "f", "t"]);
}
#[test] fn empty_associative_array_uses_assoc_entries() {
	assert_eq!(run_prints("<?php $files = []; $files['a.knt'] = '/tmp/a.knt'; echo empty($files) ? 't' : 'f'; echo count($files);"), &["f", "1"]);
}
#[test] fn associative_array_key_first_and_foreach_entries() {
	assert_eq!(run_prints("<?php $files = []; $files['a.knt'] = '/tmp/a'; $files['b.knt'] = '/tmp/b'; echo array_key_first($files); foreach ($files as $key => $path) { echo $key; echo $path; }"), &["a.knt", "a.knt", "/tmp/a", "b.knt", "/tmp/b"]);
}
#[test] fn associative_array_foreach_values() {
	assert_eq!(run_prints("<?php $files = []; $files['a.knt'] = '/tmp/a'; $files['b.knt'] = '/tmp/b'; foreach ($files as $path) { echo $path; }"), &["/tmp/a", "/tmp/b"]);
}
#[test] fn associative_array_truthiness_in_if() {
	assert_eq!(run_prints("<?php $section = ['properties' => [], 'notes' => []]; if ($section) { echo 'yes'; } else { echo 'no'; } $section['properties']['NN'] = 'Work'; if ($section) { echo 'yes'; } else { echo 'no'; }"), &["yes", "yes"]);
}
#[test] fn end_returns_last_value() {
	assert_eq!(run_prints("<?php $items = []; $items['a'] = 'one'; $items['b'] = 'two'; echo end($items);"), &["two"]);
}
#[test] fn iconv_passthrough_string() {
	assert_eq!(run_prints("<?php echo iconv('WINDOWS-1252', 'UTF-8//TRANSLIT', 'Info Base');"), &["Info Base"]);
}
#[test] fn unset_var() { compile_ok("<?php $x = 1; unset($x);"); }
