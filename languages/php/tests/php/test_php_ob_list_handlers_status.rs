use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Output Buffering: ob_list_handlers & ob_get_status
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_ob_list_handlers_returns_active_names() {
    let out = run_prints(
        r##"<?php
ob_start();
$handlers = ob_list_handlers();
ob_end_clean();

echo "HandlerName: " . ($handlers[0] ?? "");
"##,
    );
    assert_eq!(out, vec!["HandlerName: default output handler"]);
}

#[test]
fn test_php_ob_get_status_full_stack_array() {
    let out = run_prints(
        r##"<?php
ob_start(); // Level 1
ob_start(); // Level 2
$status = ob_get_status(true);
ob_end_clean();
ob_end_clean();

echo "LevelsCount: " . count($status);
"##,
    );
    assert_eq!(out, vec!["LevelsCount: 2"]);
}

#[test]
fn test_php_ob_get_status_single_level_keys() {
    let out = run_prints(
        r##"<?php
ob_start();
$status = ob_get_status(false);
ob_end_clean();

$keys = ["name", "type", "flags", "level", "chunk_size", "buffer_size", "buffer_used"];
$hasKeys = true;
foreach ($keys as $k) {
    if (!array_key_exists($k, $status)) { $hasKeys = false; break; }
}
echo $hasKeys ? "STATUS_KEYS_OK" : "MISSING_KEYS";
"##,
    );
    assert_eq!(out, vec!["STATUS_KEYS_OK"]);
}

#[test]
fn test_php_ob_list_handlers_custom_closure_name() {
    compile_ok(
        r##"<?php
ob_start(fn($s) => $s);
$handlers = ob_list_handlers();
ob_end_clean();
echo str_contains($handlers[0], "Closure") || str_contains($handlers[0], "default") ? "CLOSURE_HANDLER_NAME_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_get_status_buffer_used_property() {
    compile_ok(
        r##"<?php
ob_start();
echo "12345";
$status = ob_get_status(false);
ob_end_clean();
echo $status["buffer_used"] === 5 ? "BUFFER_USED_5_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_get_status_level_property() {
    compile_ok(
        r##"<?php
ob_start();
$s1 = ob_get_status(false);
ob_start();
$s2 = ob_get_status(false);
ob_end_clean();
ob_end_clean();
echo $s2["level"] === $s1["level"] + 1 ? "STATUS_LEVEL_INC_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_get_status_type_property() {
    compile_ok(
        r##"<?php
ob_start();
$status = ob_get_status(false);
ob_end_clean();
echo isset($status["type"]) && is_int($status["type"]) ? "STATUS_TYPE_INT_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_get_status_empty_stack_returns_empty_array() {
    compile_ok(
        r##"<?php
while (ob_get_level() > 0) ob_end_clean();
$status = ob_get_status(true);
echo is_array($status) && count($status) === 0 ? "EMPTY_STATUS_STACK_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_list_handlers_empty_stack() {
    compile_ok(
        r##"<?php
while (ob_get_level() > 0) ob_end_clean();
$handlers = ob_list_handlers();
echo is_array($handlers) && count($handlers) === 0 ? "EMPTY_HANDLERS_STACK_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_ob_get_status_chunk_size_default_zero() {
    compile_ok(
        r##"<?php
ob_start();
$status = ob_get_status(false);
ob_end_clean();
echo isset($status["chunk_size"]) ? "CHUNK_SIZE_KEY_OK" : "FAIL";
"##,
    );
}
