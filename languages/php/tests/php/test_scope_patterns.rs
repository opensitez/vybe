use super::helpers::{compile_ok, run_prints};

// ── Outer variable invisible inside function ──────────────────

#[test]
fn function_cannot_see_outer_variable() {
    compile_ok(
        r#"<?php
$secret = 42;
function noAccess(): mixed {
    return isset($secret) ? $secret : null;
}
echo noAccess() === null ? 'hidden' : 'leaked';
"#,
    );
}

// ── global keyword ────────────────────────────────────────────

#[test]
fn global_keyword_read() {
    compile_ok(
        r#"<?php
$counter = 10;
function readGlobal(): int {
    global $counter;
    return $counter;
}
echo readGlobal();
"#,
    );
}

#[test]
fn global_keyword_modify() {
    compile_ok(
        r#"<?php
$total = 0;
function addToTotal(int $n): void {
    global $total;
    $total += $n;
}
addToTotal(5);
addToTotal(3);
echo $total;
"#,
    );
}

// ── static variables ──────────────────────────────────────────

#[test]
fn static_var_retains_across_calls() {
    compile_ok(
        r#"<?php
function increment(): int {
    static $n = 0;
    return ++$n;
}
echo increment();
echo increment();
echo increment();
"#,
    );
}

#[test]
fn static_var_in_recursion() {
    compile_ok(
        r#"<?php
function depth(): int {
    static $calls = 0;
    $calls++;
    if ($calls < 4) depth();
    return $calls;
}
echo depth();
"#,
    );
}

#[test]
fn static_vars_independent_per_function() {
    compile_ok(
        r#"<?php
function counterA(): int { static $n = 0; return ++$n; }
function counterB(): int { static $n = 0; return ++$n; }
counterA(); counterA();
counterB();
echo counterA();
echo counterB();
"#,
    );
}

// ── Nested function definitions ───────────────────────────────

#[test]
fn nested_function_definition_inner_callable() {
    compile_ok(
        r#"<?php
function outer(): void {
    function inner(): string { return 'inside'; }
}
outer();
echo inner();
"#,
    );
}

#[test]
fn function_defined_inside_if_block() {
    compile_ok(
        r#"<?php
$flag = true;
if ($flag) {
    function conditionalFn(): int { return 99; }
}
echo conditionalFn();
"#,
    );
}

// ── Closure capture semantics ─────────────────────────────────

#[test]
fn closure_use_by_value_snapshot() {
    compile_ok(
        r#"<?php
$x = 1;
$fn = function() use ($x) { return $x; };
$x = 99;
echo $fn();
"#,
    );
}

#[test]
fn closure_use_by_reference_mutates_outer() {
    compile_ok(
        r#"<?php
$total = 0;
$add = function(int $n) use (&$total): void { $total += $n; };
$add(10);
$add(5);
echo $total;
"#,
    );
}

// ── Arrow function capture ────────────────────────────────────

#[test]
fn arrow_fn_auto_captures_outer_scope() {
    compile_ok(
        r#"<?php
$multiplier = 7;
$fn = fn(int $x) => $x * $multiplier;
echo $fn(6);
"#,
    );
}

#[test]
fn arrow_fn_cannot_mutate_outer_scope() {
    compile_ok(
        r#"<?php
$x = 10;
$fn = fn() => $x + 1;
$fn();
echo $x;
"#,
    );
}

// ── Nested closures sharing outer ────────────────────────────

#[test]
fn nested_closures_share_captured_ref() {
    compile_ok(
        r#"<?php
$log = [];
$push = function(string $msg) use (&$log): void { $log[] = $msg; };
$pushTwice = function(string $msg) use ($push): void { $push($msg); $push($msg); };
$pushTwice('hi');
echo count($log);
"#,
    );
}

// ── Class method scope ────────────────────────────────────────

#[test]
fn class_method_uses_this() {
    compile_ok(
        r#"<?php
class Box {
    private int $value;
    public function __construct(int $v) { $this->value = $v; }
    public function get(): int { return $this->value; }
}
echo (new Box(77))->get();
"#,
    );
}

#[test]
fn static_method_no_this() {
    compile_ok(
        r#"<?php
class MathHelper {
    public static function square(int $n): int { return $n * $n; }
}
echo MathHelper::square(9);
"#,
    );
}

// ── Variable shadowing ────────────────────────────────────────

#[test]
fn closure_local_shadows_outer_by_value() {
    compile_ok(
        r#"<?php
$name = 'outer';
$fn = function() use ($name): string {
    $name = 'inner';
    return $name;
};
echo $fn();
echo $name;
"#,
    );
}

#[test]
fn parameter_name_same_as_global() {
    compile_ok(
        r#"<?php
$x = 'global';
function echo_param(string $x): void {
    echo $x;
}
echo_param('local');
echo $x;
"#,
    );
}

// ── Class constant access ─────────────────────────────────────

#[test]
fn class_constant_from_instance_method() {
    compile_ok(
        r#"<?php
class Protocol {
    const VERSION = '1.0';
    public function version(): string { return self::VERSION; }
}
echo (new Protocol())->version();
"#,
    );
}

#[test]
fn self_vs_static_in_static_context() {
    compile_ok(
        r#"<?php
class Base {
    protected static string $label = 'Base';
    public static function selfLabel(): string  { return self::$label; }
    public static function lateLabel(): string  { return static::$label; }
}
class Child extends Base {
    protected static string $label = 'Child';
}
echo Base::selfLabel();
echo Child::lateLabel();
"#,
    );
}

// ── Constants are globally scoped ────────────────────────────

#[test]
fn constant_visible_everywhere() {
    compile_ok(
        r#"<?php
define('SITE_NAME', 'Vybe');
function getSite(): string {
    return SITE_NAME;
}
class Config {
    public function name(): string { return SITE_NAME; }
}
echo getSite();
echo (new Config())->name();
"#,
    );
}

#[test]
fn scope_global_modification_persists_across_calls_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$count = 1;
function bump_global(): int {
    global $count;
    return ++$count;
}
echo bump_global();
echo bump_global();
bump_global();
echo $count;
"#,
        ),
        vec!["234"]
    );
}

#[test]
fn scope_static_function_retains_value_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
function seq(): int {
    static $n = 0;
    return ++$n;
}
echo seq();
echo seq();
echo seq();
"#,
        ),
        vec!["123"]
    );
}

#[test]
fn scope_closure_by_ref_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$value = 1;
$inc = function(int $n) use (&$value): void { $value += $n; };
$inc(2);
$inc(3);
echo $value;
"#,
        ),
        vec!["6"]
    );
}

#[test]
fn scope_global_reference_alias_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$count = 2;
function bump_by_alias(): void {
    global $count;
    $alias = &$count;
    $alias += 3;
}
bump_by_alias();
echo $count;
"#,
        ),
        vec!["5"]
    );
}

#[test]
fn scope_global_inside_closure_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$label = 'outer';
$writer = function(string $value): void {
    global $label;
    $label = $value;
};
$writer('inner');
echo $label;
"#,
        ),
        vec!["inner"]
    );
}

#[test]
fn scope_static_recurse_counter_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
function walk(int $n): int {
    static $calls = 0;
    $calls++;
    if ($n <= 0) return $calls;
    return walk($n - 1);
}
echo walk(2);
echo walk(0);
"#,
        ),
        vec!["34"]
    );
}

#[test]
fn scope_function_name_resolution_with_variable_function_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
function dynamicScopeTarget(): string { return 'scope'; }
$fn = 'dynamicScopeTarget';
echo $fn();
"#,
        ),
        vec!["scope"]
    );
}

#[test]
fn scope_eval_defines_function_at_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
eval('function scope_eval_target(): string { return "evaled"; }');
echo scope_eval_target();
"#,
        ),
        vec!["evaled"]
    );
}

#[test]
fn scope_arrow_function_reads_outer_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$base = 10;
$double = fn(int $n) => $n * 2;
$read = fn() => $double($base);
echo $double(3);
echo '|';
echo $read();
"#,
        ),
        vec!["6|20"]
    );
}

#[test]
fn scope_global_keyword_with_shadowed_local_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$counter = 10;
function increment_global(): int {
    global $counter;
    $counter += 5;
    return $counter;
}
function read_local_counter(): int {
    $counter = 2;
    return $counter;
}
echo read_local_counter();
echo '|';
echo increment_global();
echo '|';
echo increment_global();
"#,
        ),
        vec!["2|15|20"]
    );
}

#[test]
fn scope_static_var_isolation_between_references_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
function first(): int {
    static $n = 1;
    $n++;
    return $n;
}
function second(): int {
    static $n = 100;
    $n += 2;
    return $n;
}
echo first();
echo '|';
echo first();
echo '|';
echo second();
"#,
        ),
        vec!["2|3|102"]
    );
}

#[test]
fn scope_closure_parameter_shadowing_outer_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$value = 'outer';
$fn = function(string $arg) use (&$value): void {
    $value = 'inner';
};
$fn('ignored');
echo $value;
"#,
        ),
        vec!["outer"]
    );
}

#[test]
fn scope_dynamic_variable_name_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$name = 'color';
$color = 'blue';
$$name = 'green';
echo $$name;
echo '|';
echo $color;
"#,
        ),
        vec!["green|blue"]
    );
}

#[test]
fn scope_variable_variables_in_nested_scope_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$prefix = 'ns';
function make_value(): void {
    $local = 'secret';
    $k = 'local';
    $prefix = 'local_prefix';
    echo $$k;
    echo '|';
    echo ${$k};
}
make_value();
echo 'done';
"#,
        ),
        vec!["secret|secret|done"]
    );
}

#[test]
fn scope_nested_function_definition_visibility_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
function define_inner(): void {
    $flag = true;
    if ($flag) {
        function nested_scoped_fn(): string { return 'ok'; }
    }
}
define_inner();
echo function_exists('nested_scoped_fn') ? nested_scoped_fn() : 'missing';
"#,
        ),
        vec!["ok"]
    );
}

#[test]
fn scope_static_property_access_from_closure_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
class Counter {
    public static int $count = 0;
}
$inc = function(): int {
    return ++Counter::$count;
};
echo $inc();
echo '|';
echo $inc();
echo '|';
echo Counter::$count;
"#,
        ),
        vec!["1|2|2"]
    );
}
