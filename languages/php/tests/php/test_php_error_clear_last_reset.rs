use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Error State: error_clear_last & error_get_last Reset
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_error_clear_last_resets_error_get_last() {
    let out = run_prints(
        r##"<?php
@trigger_error("Temporary notice", E_USER_NOTICE);
$before = error_get_last();

error_clear_last();
$after = error_get_last();

echo "Before=" . ($before !== null ? "SET" : "NULL") . " After=" . ($after === null ? "NULL" : "SET");
"##,
    );
    assert_eq!(out, vec!["Before=SET After=NULL"]);
}

#[test]
fn test_php_error_get_last_structure_keys() {
    let out = run_prints(
        r##"<?php
@trigger_error("Structured error", E_USER_WARNING);
$err = error_get_last();
error_clear_last();

echo "Type={$err['type']} Message={$err['message']}";
"##,
    );
    assert_eq!(out, vec!["Type=512 Message=Structured error"]);
}

#[test]
fn test_php_error_clear_last_when_no_error_exists() {
    compile_ok(
        r##"<?php
error_clear_last();
$err = error_get_last();
echo $err === null ? "NO_ERROR_NULL_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_error_get_last_file_and_line_properties() {
    compile_ok(
        r##"<?php
@trigger_error("File line test", E_USER_NOTICE);
$err = error_get_last();
error_clear_last();
echo isset($err["file"]) && isset($err["line"]) && $err["line"] > 0 ? "FILE_LINE_KEYS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_error_clear_last_between_multiple_errors() {
    compile_ok(
        r##"<?php
@trigger_error("First notice", E_USER_NOTICE);
error_clear_last();
@trigger_error("Second notice", E_USER_NOTICE);
$err = error_get_last();
error_clear_last();
echo $err["message"] === "Second notice" ? "SECOND_NOTICE_CAPTURED_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_error_get_last_user_error_type() {
    compile_ok(
        r##"<?php
@trigger_error("User Error Level", E_USER_ERROR);
$err = error_get_last();
error_clear_last();
echo $err["type"] === E_USER_ERROR ? "E_USER_ERROR_MATCH_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_error_get_last_deprecated_type() {
    compile_ok(
        r##"<?php
@trigger_error("Deprecation warning", E_USER_DEPRECATED);
$err = error_get_last();
error_clear_last();
echo $err["type"] === E_USER_DEPRECATED ? "E_USER_DEPRECATED_MATCH_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_error_clear_last_inside_error_handler() {
    compile_ok(
        r##"<?php
set_error_handler(function($errno, $errstr) {
    error_clear_last();
    return true;
});
@trigger_error("Inside handler", E_USER_NOTICE);
restore_error_handler();
echo error_get_last() === null ? "CLEAR_INSIDE_HANDLER_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_error_get_last_returns_null_after_clear() {
    compile_ok(
        r##"<?php
@trigger_error("Test", E_USER_NOTICE);
error_clear_last();
echo error_get_last() === null ? "CLEAR_NULL_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_error_clear_last_multiple_sequential_clears() {
    compile_ok(
        r##"<?php
error_clear_last();
error_clear_last();
error_clear_last();
echo error_get_last() === null ? "SEQUENTIAL_CLEARS_OK" : "FAIL";
"##,
    );
}
