use super::helpers::compile_ok;

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
