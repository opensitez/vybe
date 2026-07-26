use super::helpers::run_prints;

// ── global keyword ────────────────────────────────────────────

#[test]
fn global_var_accessible_in_function() {
    assert_eq!(
        run_prints(
            r#"<?php
$counter = 0;
function increment(): void { global $counter; $counter++; }
increment(); increment(); increment();
echo $counter;
"#
        ),
        vec!["3"]
    );
}
#[test]
fn global_multiple_vars() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = 10; $b = 20;
function sum_globals(): int { global $a, $b; return $a + $b; }
echo sum_globals();
"#
        ),
        vec!["30"]
    );
}
#[test]
fn global_write_persists() {
    assert_eq!(
        run_prints(
            r#"<?php
$msg = 'hello';
function modify(): void { global $msg; $msg = 'world'; }
modify();
echo $msg;
"#
        ),
        vec!["world"]
    );
}
#[test]
fn function_scope_isolates_from_global() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = 99;
function noAccess(): string { return isset($x) ? 'visible' : 'hidden'; }
echo noAccess();
"#
        ),
        vec!["hidden"]
    );
}

// ── static variables ──────────────────────────────────────────

#[test]
fn static_var_persists_between_calls() {
    assert_eq!(
        run_prints(
            r#"<?php
function counter(): int { static $n = 0; return ++$n; }
echo counter() . ',' . counter() . ',' . counter();
"#
        ),
        vec!["1,2,3"]
    );
}
#[test]
fn static_var_initialized_once() {
    assert_eq!(
        run_prints(
            r#"<?php
function once(): int { static $x = 100; $x += 5; return $x; }
echo once() . ',' . once();
"#
        ),
        vec!["105,110"]
    );
}
#[test]
fn static_var_in_recursive_function() {
    assert_eq!(
        run_prints(
            r#"<?php
function depth(): int {
    static $d = 0;
    $d++;
    if ($d < 3) depth();
    return $d;
}
echo depth();
"#
        ),
        vec!["3"]
    );
}
#[test]
fn static_var_per_function_scope() {
    assert_eq!(
        run_prints(
            r#"<?php
function fn1(): int { static $n = 0; return ++$n; }
function fn2(): int { static $n = 0; return ++$n; }
fn1(); fn1();
fn2();
echo fn1() . ',' . fn2();
"#
        ),
        vec!["3,2"]
    );
}

// ── Static class properties ───────────────────────────────────

#[test]
fn static_class_property_shared() {
    assert_eq!(
        run_prints(
            r#"<?php
class Registry {
    public static int $count = 0;
    public function __construct() { self::$count++; }
}
new Registry; new Registry; new Registry;
echo Registry::$count;
"#
        ),
        vec!["3"]
    );
}
#[test]
fn static_class_method_without_instance() {
    assert_eq!(
        run_prints(
            r#"<?php
class MathUtil { public static function cube(int $n): int { return $n ** 3; } }
echo MathUtil::cube(4);
"#
        ),
        vec!["64"]
    );
}
#[test]
fn static_property_inheritance_not_shared() {
    assert_eq!(
        run_prints(
            r#"<?php
class Base { public static int $val = 0; }
class ChildA extends Base {}
class ChildB extends Base {}
ChildA::$val = 10;
echo Base::$val . ',' . ChildA::$val . ',' . ChildB::$val;
"#
        ),
        vec!["10,10,10"]
    );
}

// ── Constants ─────────────────────────────────────────────────

#[test]
fn define_constant() {
    assert_eq!(
        run_prints(r#"<?php define('APP_NAME', 'Vybe'); echo APP_NAME; "#),
        vec!["Vybe"]
    );
}
#[test]
fn class_constant_access() {
    assert_eq!(
        run_prints(
            r#"<?php
class Color { const RED = 'red'; const GREEN = 'green'; }
echo Color::RED . ',' . Color::GREEN;
"#
        ),
        vec!["red,green"]
    );
}
#[test]
fn interface_constant_access() {
    assert_eq!(
        run_prints(
            r#"<?php
interface Status { const ACTIVE = 1; const INACTIVE = 0; }
class User implements Status {}
echo User::ACTIVE . ',' . User::INACTIVE;
"#
        ),
        vec!["1,0"]
    );
}
#[test]
fn constant_in_expression() {
    assert_eq!(
        run_prints(
            r#"<?php
define('TAX_RATE', 0.2);
$price = 100;
echo $price * (1 + TAX_RATE);
"#
        ),
        vec!["120"]
    );
}

// ── Nested function scope ─────────────────────────────────────

#[test]
fn nested_function_defined_on_call() {
    assert_eq!(
        run_prints(
            r#"<?php
function outer(): void {
    function inner(): string { return 'inner'; }
}
outer();
echo inner();
"#
        ),
        vec!["inner"]
    );
}
#[test]
fn closure_vs_function_scope() {
    assert_eq!(
        run_prints(
            r#"<?php
$x = 5;
$fn = function() use ($x) { return $x * 2; };
$x = 10;
echo $fn();
"#
        ),
        vec!["10"]
    );
}
#[test]
fn closure_use_by_reference() {
    assert_eq!(
        run_prints(
            r#"<?php
$total = 0;
$add = function(int $n) use (&$total): void { $total += $n; };
$add(3); $add(7); $add(10);
echo $total;
"#
        ),
        vec!["20"]
    );
}

#[test]
fn global_reference_alias_in_nested_function() {
    assert_eq!(
        run_prints(
            r#"<?php
$seed = 1;
function set_seed(int $v): void {
    global $seed;
    $r = &$seed;
    $r = $v;
}
set_seed(7);
echo $seed;
"#
        ),
        vec!["7"]
    );
}

#[test]
fn global_in_closure_with_reference() {
    assert_eq!(
        run_prints(
            r#"<?php
$token = 'init';
$writer = function(string $value): void {
    global $token;
    $token = $value;
};
$writer('done');
echo $token;
"#
        ),
        vec!["done"]
    );
}

#[test]
fn static_property_isolation_in_methods() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    public function inc(): int { static $n = 0; return ++$n; }
}
$a = new Counter();
$b = new Counter();
echo $a->inc() . $a->inc();
echo '|';
echo $b->inc();
"#
        ),
        vec!["12|3"]
    );
}

#[test]
fn static_in_multiple_functions_with_same_name() {
    assert_eq!(
        run_prints(
            r#"<?php
function first(): int { static $n = 0; return ++$n; }
function second(): int { static $n = 5; return ++$n; }
echo first() . '|' . first() . '|' . second();
"#
        ),
        vec!["1|2|6"]
    );
}

crate::php_cases! {
    static_local_variable_increments_per_call => {
        r#"<?php
function counter(): int { static $n = 0; return ++$n; }
echo counter() . counter();
"#,
        ["12"]
    };
}
