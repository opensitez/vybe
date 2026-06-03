use super::helpers::compile_ok;

// ── isset on chained array key access ────────────────────────────
#[test]
fn isset_chained_array_key() {
    compile_ok(
        r#"<?php
$config = ['db' => ['host' => 'localhost', 'port' => 3306]];
echo isset($config['db']['host']) ? 'set' : 'missing';
echo isset($config['db']['password']) ? 'set' : 'missing';
echo isset($config['cache']['host']) ? 'set' : 'missing';
"#,
    );
}

// ── isset returns false for null-valued key ───────────────────────
#[test]
fn isset_false_for_null_value() {
    compile_ok(
        r#"<?php
$a = null;
$b = 0;
$c = '';
echo isset($a) ? 'set' : 'unset';
echo isset($b) ? 'set' : 'unset';
echo isset($c) ? 'set' : 'unset';
"#,
    );
}

// ── unset array element and verify after loop ─────────────────────
#[test]
fn unset_array_element_after_loop() {
    compile_ok(
        r#"<?php
$items = ['a', 'b', 'c', 'd'];
$toRemove = [];
foreach ($items as $k => $v) {
    if ($v === 'b' || $v === 'd') {
        $toRemove[] = $k;
    }
}
foreach ($toRemove as $k) {
    unset($items[$k]);
}
echo implode(',', array_values($items));
"#,
    );
}

// ── empty() on zero, empty string, empty array, null, false, "0" ─
#[test]
fn empty_on_falsy_values() {
    compile_ok(
        r#"<?php
$checks = [0, '', [], null, false, '0', 1, 'x', [1], true];
foreach ($checks as $v) {
    echo empty($v) ? '1' : '0';
}
"#,
    );
}

// ── defined() checking a constant ────────────────────────────────
#[test]
fn defined_check_constant() {
    compile_ok(
        r#"<?php
define('APP_VERSION', '1.0.0');
echo defined('APP_VERSION') ? 'yes' : 'no';
echo defined('MISSING_CONST') ? 'yes' : 'no';
"#,
    );
}

// ── constant() getting value by name ─────────────────────────────
#[test]
fn constant_get_value() {
    compile_ok(
        r#"<?php
define('MAX_RETRIES', 3);
$name = 'MAX_RETRIES';
echo constant($name);
"#,
    );
}

// ── define() with boolean value ───────────────────────────────────
#[test]
fn define_boolean_value() {
    compile_ok(
        r#"<?php
define('DEBUG_MODE', false);
define('FEATURE_FLAG', true);
echo DEBUG_MODE ? 'debug' : 'prod';
echo FEATURE_FLAG ? 'on' : 'off';
"#,
    );
}

// ── variable variable $$name ──────────────────────────────────────
#[test]
fn variable_variable_basic_assign() {
    compile_ok(
        r#"<?php
$varName = 'color';
$$varName = 'blue';
echo $color;
echo $$varName;
"#,
    );
}

// ── variable variable in array access ────────────────────────────
#[test]
fn variable_variable_array_access() {
    compile_ok(
        r#"<?php
$fields = ['x', 'y', 'z'];
$x = 10;
$y = 20;
$z = 30;
$sum = 0;
foreach ($fields as $name) {
    $sum += $$name;
}
echo $sum;
"#,
    );
}

// ── get_defined_vars returns array ───────────────────────────────
#[test]
fn get_defined_vars_returns_array() {
    compile_ok(
        r#"<?php
$alpha = 1;
$beta  = 2;
$vars = get_defined_vars();
echo is_array($vars) ? 'yes' : 'no';
"#,
    );
}

// ── static variable retains value across calls ────────────────────
#[test]
fn static_var_retains_value() {
    compile_ok(
        r#"<?php
function counter(): int {
    static $count = 0;
    $count++;
    return $count;
}
echo counter();
echo counter();
echo counter();
"#,
    );
}

// ── static variable initialized only once ────────────────────────
#[test]
fn static_var_initialized_once() {
    compile_ok(
        r#"<?php
function makeId(): string {
    static $id = 'ID-000';
    return $id;
}
echo makeId();
echo makeId();
"#,
    );
}

// ── global variable modified inside function ──────────────────────
#[test]
fn global_var_modified_in_function() {
    compile_ok(
        r#"<?php
$score = 0;
function addPoints(int $pts): void {
    global $score;
    $score += $pts;
}
addPoints(10);
addPoints(5);
echo $score;
"#,
    );
}

// ── global array modified inside function ────────────────────────
#[test]
fn global_array_modified_in_function() {
    compile_ok(
        r#"<?php
$log = [];
function logEvent(string $msg): void {
    global $log;
    $log[] = $msg;
}
logEvent('start');
logEvent('stop');
echo implode(',', $log);
"#,
    );
}

// ── list() swap two variables ─────────────────────────────────────
#[test]
fn list_swap_variables() {
    compile_ok(
        r#"<?php
$a = 'first';
$b = 'second';
[$a, $b] = [$b, $a];
echo $a;
echo $b;
"#,
    );
}

// ── list() destructure function return ───────────────────────────
#[test]
fn list_from_function_return() {
    compile_ok(
        r#"<?php
function minMax(array $arr): array {
    return [min($arr), max($arr)];
}
[$lo, $hi] = minMax([3, 1, 4, 1, 5, 9]);
echo $lo;
echo $hi;
"#,
    );
}

// ── extract() with EXTR_PREFIX_ALL to avoid conflicts ────────────
#[test]
fn extract_with_prefix() {
    compile_ok(
        r#"<?php
$name = 'outer';
$data = ['name' => 'inner', 'age' => 25];
extract($data, EXTR_PREFIX_ALL, 'row');
echo $name;
echo $row_name;
echo $row_age;
"#,
    );
}

// ── compact() with array of variable names ────────────────────────
#[test]
fn compact_variable_names() {
    compile_ok(
        r#"<?php
$city    = 'Paris';
$country = 'France';
$pop     = 2161000;
$result = compact('city', 'country', 'pop');
echo $result['city'];
echo $result['country'];
"#,
    );
}

// ── isset() on undefined property of object ───────────────────────
#[test]
fn isset_undefined_object_property() {
    compile_ok(
        r#"<?php
class Config {
    public string $host = 'localhost';
}
$c = new Config();
echo isset($c->host)    ? 'set' : 'unset';
echo isset($c->missing) ? 'set' : 'unset';
"#,
    );
}

// ── unset() on object property ────────────────────────────────────
#[test]
fn unset_object_property() {
    compile_ok(
        r#"<?php
class Bag {
    public string $item = 'apple';
    public int    $qty  = 5;
}
$b = new Bag();
echo isset($b->item) ? 'set' : 'gone';
unset($b->item);
echo isset($b->item) ? 'set' : 'gone';
echo $b->qty;
"#,
    );
}
