use super::helpers::compile_ok;

// ── Integer limits ────────────────────────────────────────────

#[test]
fn php_int_max() {
    compile_ok(
        r#"<?php
echo PHP_INT_MAX;
"#,
    );
}

#[test]
fn php_int_min() {
    compile_ok(
        r#"<?php
echo PHP_INT_MIN;
"#,
    );
}

#[test]
fn php_int_size() {
    compile_ok(
        r#"<?php
echo PHP_INT_SIZE;
"#,
    );
}

// ── Float limits ──────────────────────────────────────────────

#[test]
fn php_float_max() {
    compile_ok(
        r#"<?php
$x = PHP_FLOAT_MAX;
echo ($x > 0) ? 'positive' : 'unexpected';
"#,
    );
}

#[test]
fn php_float_min() {
    compile_ok(
        r#"<?php
$x = PHP_FLOAT_MIN;
echo ($x > 0) ? 'positive' : 'unexpected';
"#,
    );
}

#[test]
fn php_float_epsilon() {
    compile_ok(
        r#"<?php
$eps = PHP_FLOAT_EPSILON;
echo ($eps > 0) ? 'positive' : 'unexpected';
"#,
    );
}

// ── Environment constants ─────────────────────────────────────

#[test]
fn php_eol_constant() {
    compile_ok(
        r#"<?php
$line = 'hello' . PHP_EOL;
echo strlen($line) > 5 ? 'has_eol' : 'no_eol';
"#,
    );
}

#[test]
fn php_major_version() {
    compile_ok(
        r#"<?php
$v = PHP_MAJOR_VERSION;
"#,
    );
}

#[test]
fn php_os_family() {
    compile_ok(
        r#"<?php
$os = PHP_OS_FAMILY;
"#,
    );
}

#[test]
fn php_sapi() {
    compile_ok(
        r#"<?php
$sapi = PHP_SAPI;
"#,
    );
}

// ── Boolean / null constants ──────────────────────────────────

#[test]
fn uppercase_true_false_null() {
    compile_ok(
        r#"<?php
$t = TRUE;
$f = FALSE;
$n = NULL;
echo $t ? 'yes' : 'no';
echo $f ? 'yes' : 'no';
echo is_null($n) ? 'null' : 'not null';
"#,
    );
}

// ── Sort constants ────────────────────────────────────────────

#[test]
fn sort_flag_constants() {
    compile_ok(
        r#"<?php
$a = [3, 1, 2];
sort($a, SORT_REGULAR);
$b = ['10', '9', '2'];
sort($b, SORT_NUMERIC);
$c = ['banana', 'apple', 'cherry'];
sort($c, SORT_STRING);
echo count($a) + count($b) + count($c);
"#,
    );
}

// ── String pad constants ──────────────────────────────────────

#[test]
fn str_pad_direction_constants() {
    compile_ok(
        r#"<?php
$left  = str_pad('5', 3, '0', STR_PAD_LEFT);
$right = str_pad('5', 3, '0', STR_PAD_RIGHT);
$both  = str_pad('5', 5, '-', STR_PAD_BOTH);
echo $left . $right . $both;
"#,
    );
}

// ── Array filter constants ────────────────────────────────────

#[test]
fn array_filter_use_key_constant() {
    compile_ok(
        r#"<?php
$arr = ['a' => 1, 'b' => 2, 'c' => 3];
$result = array_filter($arr, fn($k) => $k !== 'b', ARRAY_FILTER_USE_KEY);
echo count($result);
"#,
    );
}

#[test]
fn array_filter_use_both_constant() {
    compile_ok(
        r#"<?php
$arr = ['x' => 10, 'y' => 5, 'z' => 20];
$result = array_filter($arr, fn($v, $k) => $k !== 'y' && $v > 8, ARRAY_FILTER_USE_BOTH);
echo count($result);
"#,
    );
}

// ── Path constants ────────────────────────────────────────────

#[test]
fn directory_separator_constant() {
    compile_ok(
        r#"<?php
$path = 'usr' . DIRECTORY_SEPARATOR . 'local' . DIRECTORY_SEPARATOR . 'bin';
echo strlen($path) > 0 ? 'ok' : 'empty';
"#,
    );
}

#[test]
fn path_separator_constant() {
    compile_ok(
        r#"<?php
$env = '/usr/bin' . PATH_SEPARATOR . '/usr/local/bin';
echo strlen($env) > 0 ? 'ok' : 'empty';
"#,
    );
}

// ── User-defined constants ────────────────────────────────────

#[test]
fn define_scalar_constant() {
    compile_ok(
        r#"<?php
define('APP_VERSION', '2.0.1');
echo APP_VERSION;
"#,
    );
}

#[test]
fn define_array_constant() {
    compile_ok(
        r#"<?php
define('ALLOWED_ROLES', ['admin', 'editor', 'viewer']);
echo count(ALLOWED_ROLES);
"#,
    );
}

#[test]
fn class_const_vs_define() {
    compile_ok(
        r#"<?php
define('GLOBAL_LIMIT', 100);
class Config {
    const LIMIT = 200;
}
echo GLOBAL_LIMIT;
echo Config::LIMIT;
"#,
    );
}

#[test]
fn constant_function_by_name() {
    compile_ok(
        r#"<?php
define('TIMEOUT', 30);
$name = 'TIMEOUT';
echo constant($name);
"#,
    );
}
