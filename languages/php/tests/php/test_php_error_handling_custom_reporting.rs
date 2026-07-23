use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Error Handling & Custom Reporting — set_error_handler, set_exception_handler, error_reporting, trigger_error, error_get_last
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_custom_error_handler_interception() {
    let out = run_prints(
        r#"<?php
$captured = [];
set_error_handler(function($errno, $errstr) use (&$captured) {
    $captured[] = "ERR[$errno]: $errstr";
    return true; // suppress default PHP handler
});

trigger_error("Custom Warning Message", E_USER_WARNING);
restore_error_handler();

echo implode(", ", $captured);
"#,
    );
    assert_eq!(out, vec!["ERR[512]: Custom Warning Message"]);
}

#[test]
fn test_php_error_reporting_bitwise_masks() {
    let out = run_prints(
        r#"<?php
$old = error_reporting(E_ALL & ~E_NOTICE);
echo (error_reporting() & E_NOTICE) === 0 ? "NOTICE_DISABLED" : "NOTICE_ENABLED";
error_reporting($old);
"#,
    );
    assert_eq!(out, vec!["NOTICE_DISABLED"]);
}

#[test]
fn test_php_error_clear_last_reset() {
    let out = run_prints(
        r#"<?php
@trigger_error("Test notice", E_USER_NOTICE);
$err1 = error_get_last();
error_clear_last();
$err2 = error_get_last();

echo ($err1 !== null ? "HAD_ERROR" : "NO") . " | " . ($err2 === null ? "CLEARED" : "NOT_CLEARED");
"#,
    );
    assert_eq!(out, vec!["HAD_ERROR | CLEARED"]);
}

#[test]
fn test_php_custom_uncaught_exception_handler() {
    let out = run_prints(
        r#"<?php
set_exception_handler(function(Throwable $e) {
    echo "UNCAUGHT: " . $e->getMessage();
});

throw new Exception("Unhandled Exception");
"#,
    );
    assert_eq!(out, vec!["UNCAUGHT: Unhandled Exception"]);
}

#[test]
fn test_php_trigger_error_user_levels() {
    compile_ok(
        r#"<?php
@trigger_error("User error message", E_USER_ERROR);
@trigger_error("User notice message", E_USER_NOTICE);
@trigger_error("User deprecated message", E_USER_DEPRECATED);
"#,
    );
}

#[test]
fn test_php_error_handler_return_false_pass_through() {
    compile_ok(
        r#"<?php
set_error_handler(function($errno, $errstr) {
    return false; // pass to default PHP error handler
});
@trigger_error("Pass through notice", E_USER_NOTICE);
restore_error_handler();
"#,
    );
}

#[test]
fn test_php_exception_handler_chaining() {
    compile_ok(
        r#"<?php
$prev = set_exception_handler(fn($e) => echo "Handler 1");
$prev2 = set_exception_handler(fn($e) => echo "Handler 2");
restore_exception_handler();
"#,
    );
}

#[test]
fn test_php_silence_operator_suppress_notices() {
    compile_ok(
        r#"<?php
$arr = [];
$val = @$arr["non_existent_key"];
echo $val === null ? "NULL_SUPPRESSED" : "FAIL";
"#,
    );
}

#[test]
fn test_php_error_log_custom_destination() {
    compile_ok(
        r#"<?php
error_log("Test error log entry", 0);
"#,
    );
}

#[test]
fn test_php_register_shutdown_function_cleanup() {
    compile_ok(
        r#"<?php
register_shutdown_function(function() {
    // Shutdown cleanup callback
});
"#,
    );
}
