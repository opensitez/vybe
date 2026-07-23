use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Error Handlers: set_error_handler, levels mask & bypass
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_set_error_handler_level_mask_filtering() {
    let out = run_prints(
        r##"<?php
$captured = [];
set_error_handler(function($errno, $errstr) use (&$captured) {
    $captured[] = "Err:$errno Str:$errstr";
}, E_USER_WARNING);

@trigger_error("Ignored notice", E_USER_NOTICE);
@trigger_error("Captured warning", E_USER_WARNING);

restore_error_handler();

echo implode("; ", $captured);
"##,
    );
    assert_eq!(out, vec!["Err:512 Str:Captured warning"]);
}

#[test]
fn test_php_set_error_handler_return_true_bypasses_internal_handler() {
    let out = run_prints(
        r##"<?php
set_error_handler(function($errno, $errstr) {
    echo "CustomHandler: $errstr";
    return true; // Bypass standard PHP error handler
});

trigger_error("Bypassed error message", E_USER_NOTICE);
restore_error_handler();
"##,
    );
    assert_eq!(out, vec!["CustomHandler: Bypassed error message"]);
}

#[test]
fn test_php_set_error_handler_return_false_bubbles_error() {
    let out = run_prints(
        r##"<?php
set_error_handler(function($errno, $errstr) {
    echo "HandlerSeen ";
    return false; // Bubble to standard handler or error suppression
});

@trigger_error("Bubbled notice", E_USER_NOTICE);
restore_error_handler();
"##,
    );
    assert_eq!(out, vec!["HandlerSeen"]);
}

#[test]
fn test_php_set_error_handler_stack_restore() {
    compile_ok(
        r##"<?php
$h1 = fn() => true;
$h2 = fn() => true;
set_error_handler($h1);
set_error_handler($h2);
$prev = restore_error_handler();
restore_error_handler();
echo "ERROR_HANDLER_STACK_OK";
"##,
    );
}

#[test]
fn test_php_error_reporting_level_mask_get_set() {
    compile_ok(
        r##"<?php
$old = error_reporting(E_ALL & ~E_NOTICE);
$current = error_reporting();
error_reporting($old);
echo $current === (E_ALL & ~E_NOTICE) ? "ERROR_REPORTING_MASK_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_set_error_handler_receives_file_and_line() {
    compile_ok(
        r##"<?php
$capturedLine = 0;
set_error_handler(function($errno, $errstr, $errfile, $errline) use (&$capturedLine) {
    $capturedLine = $errline;
    return true;
});
@trigger_error("Test line capture", E_USER_NOTICE);
restore_error_handler();
echo $capturedLine > 0 ? "LINE_CAPTURE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_set_error_handler_null_resets_to_builtin() {
    compile_ok(
        r##"<?php
set_error_handler(fn() => true);
set_error_handler(null);
echo "NULL_RESET_OK";
"##,
    );
}

#[test]
fn test_php_set_error_handler_e_all_mask() {
    compile_ok(
        r##"<?php
$count = 0;
set_error_handler(function() use (&$count) { $count++; return true; }, E_ALL);
@trigger_error("Notice 1", E_USER_NOTICE);
@trigger_error("Warning 1", E_USER_WARNING);
restore_error_handler();
echo $count === 2 ? "E_ALL_MASK_2_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_set_error_handler_class_method_array_callable() {
    compile_ok(
        r##"<?php
class ErrorCatcher {
    public static function handle($errno, $errstr): bool { return true; }
}
set_error_handler([ErrorCatcher::class, "handle"]);
@trigger_error("Class handler test", E_USER_NOTICE);
restore_error_handler();
echo "CLASS_METHOD_HANDLER_OK";
"##,
    );
}

#[test]
fn test_php_set_error_handler_deprecated_level() {
    compile_ok(
        r##"<?php
$dep = false;
set_error_handler(function($errno) use (&$dep) {
    if ($errno === E_USER_DEPRECATED) $dep = true;
    return true;
});
@trigger_error("Deprecated feature used", E_USER_DEPRECATED);
restore_error_handler();
echo $dep ? "E_USER_DEPRECATED_OK" : "FAIL";
"##,
    );
}
