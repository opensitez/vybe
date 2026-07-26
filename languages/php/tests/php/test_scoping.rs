use super::helpers::{compile_ok, run_prints};

// ── Variable scoping ────────────────────────────────────────
#[test]
fn global_var() {
    compile_ok("<?php $x = 42; echo $x;");
}
#[test]
fn function_scope_isolation() {
    compile_ok("<?php $x = 1; function foo() { $x = 2; return $x; } echo foo(); echo $x;");
}
#[test]
fn global_keyword() {
    compile_ok("<?php $x = 10; function foo() { global $x; return $x + 1; } echo foo();");
}
#[test]
fn nested_function_scope() {
    compile_ok(
        "<?php function outer() { $x = 1; function inner() { return 42; } return inner() + $x; } echo outer();",
    );
}
#[test]
fn block_scope() {
    compile_ok("<?php if (true) { $x = 1; } echo $x;");
} // PHP: block doesn't create scope
#[test]
fn for_scope() {
    compile_ok("<?php for ($i = 0; $i < 5; $i++) {} echo $i;");
} // $i visible after loop
#[test]
fn static_var() {
    compile_ok(
        "<?php function counter() { $c = 0; $c++; return $c; } echo counter(); echo counter();",
    );
}
#[test]
fn closure_scope() {
    compile_ok(
        "<?php $x = 'outer'; $fn = function() { $x = 'inner'; return $x; }; echo $fn(); echo $x;",
    );
}
#[test]
fn closure_use_scope() {
    compile_ok(
        "<?php $x = 'hello'; $fn = function() use ($x) { return $x; }; $x = 'changed'; echo $fn();",
    );
}
#[test]
fn arrow_fn_scope() {
    compile_ok("<?php $x = 5; $fn = fn() => $x; echo $fn();");
}

// ── Variable variables (compile only) ───────────────────────
#[test]
fn multiple_assignment() {
    compile_ok("<?php $a = $b = $c = 0; echo $a + $b + $c;");
}
#[test]
fn swap_vars() {
    compile_ok("<?php $a = 1; $b = 2; $tmp = $a; $a = $b; $b = $tmp; echo $a . $b;");
}

// ── Constants ───────────────────────────────────────────────
#[test]
fn const_global() {
    compile_ok("<?php const PI = 3.14159; echo PI;");
}
#[test]
fn define_constant() {
    compile_ok("<?php define('MAX_SIZE', 100); echo MAX_SIZE;");
}
#[test]
fn class_constants() {
    compile_ok(
        "<?php class Config { const DB = 'mysql'; const PORT = 3306; } echo Config::DB . ':' . Config::PORT;",
    );
}

#[test]
fn static_property_scope() {
    compile_ok(
        "<?php class Counter { public static int $count = 0; public static function next(): int { return self::$count++; } }\nCounter::$count = 5;\necho Counter::next();\necho '|';\necho Counter::next();",
    );
}

#[test]
fn closure_by_reference_capture() {
    compile_ok(
        "<?php $value = 1; $inc = function() use (&$value) { $value++; }; $inc(); $inc(); echo $value;",
    );
}

#[test]
fn nested_closure_capture_from_arrow_function() {
    compile_ok(
        "<?php $factor = 2; $double = fn(int $n) => $n * $factor; $twice = function(int $n) use ($double): int { return $double($n); }; echo $twice(4);",
    );
}

#[test]
fn static_vars_in_functions_isolated() {
    compile_ok(
        "<?php function next_id(): int { static $id = 0; return ++$id; } echo next_id(); echo '-'; echo next_id();",
    );
}

#[test]
fn dynamic_variable_reference_and_functions() {
    compile_ok(
        "<?php $name = 'target'; $fn = function() use ($name) { return $name; }; echo $fn();\n$var = 'name'; echo '|' . $$var;",
    );
}

#[test]
fn global_scope_runtime_bridge() {
    let out = run_prints(
        r#"<?php
$count = 10;
function increment_global(): int {
    global $count;
    return ++$count;
}
echo $count . '|';
echo increment_global() . '|';
echo $count;
"#,
    );
    assert_eq!(out, vec!["10|11|11"]);
}

#[test]
fn closure_capture_by_value_and_by_reference_runtime() {
    let out = run_prints(
        r#"<?php
$base = 5;
$value_copy = $base;
$value_ref = &$base;
$f = function() use ($value_copy) { return $value_copy + 1; };
$g = function() use (&$value_ref) { return ++$value_ref; };
echo $base . '|';
echo $f() . '|';
$base = 8;
echo $g() . '|';
echo $base;
"#,
    );
    assert_eq!(out, vec!["5|6|9|9"]);
}

#[test]
fn static_scope_retains_between_calls_runtime() {
    let out = run_prints(
        r#"<?php
function stepper(): int {
    static $counter = 0;
    return ++$counter;
}
echo stepper() . '|' . stepper() . '|' . stepper();
"#,
    );
    assert_eq!(out, vec!["1|2|3"]);
}

#[test]
fn nested_function_visibility_runtime() {
    let out = run_prints(
        r#"<?php
$prefix = 'outer';
function outer_scope(): string {
    function inner_scope_fn(): string { return 'inner'; }
    return inner_scope_fn();
}
echo outer_scope() . '|' . inner_scope_fn();
"#,
    );
    assert_eq!(out, vec!["inner|inner"]);
}

#[test]
fn for_loop_index_scoping_runtime() {
    let out = run_prints(
        r#"<?php
$sum = 0;
for ($i = 0; $i < 3; $i++) {
    $sum += $i;
}
echo $sum . '|' . $i;
"#,
    );
    assert_eq!(out, vec!["3|3"]);
}
