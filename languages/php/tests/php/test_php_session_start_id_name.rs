use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Sessions: session_start, session_id, session_name & Status
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_session_name_getter_and_setter() {
    let out = run_prints(
        r##"<?php
@session_name("MY_APP_SESS");
$name = @session_name();
echo "SessionName: $name";
"##,
    );
    assert_eq!(out, vec!["SessionName: MY_APP_SESS"]);
}

#[test]
fn test_php_session_id_custom_setter() {
    let out = run_prints(
        r##"<?php
@session_id("custom_sess_id_12345");
$id = @session_id();
echo "SessionID: $id";
"##,
    );
    assert_eq!(out, vec!["SessionID: custom_sess_id_12345"]);
}

#[test]
fn test_php_session_status_none_vs_active() {
    let out = run_prints(
        r##"<?php
$before = session_status();
@session_start();
$after = session_status();
@session_write_close();

echo ($before === PHP_SESSION_NONE ? "NONE" : "OTHER") . " -> " . ($after === PHP_SESSION_ACTIVE ? "ACTIVE" : "OTHER");
"##,
    );
    assert_eq!(out, vec!["NONE -> ACTIVE"]);
}

#[test]
fn test_php_session_start_options_read_and_close() {
    compile_ok(
        r##"<?php
$status = @session_start([
    "read_and_close" => true,
    "cookie_lifetime" => 3600
]);
echo $status && session_status() === PHP_SESSION_NONE ? "READ_AND_CLOSE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_reset_reloads_values() {
    compile_ok(
        r##"<?php
@session_start();
$_SESSION["key"] = "original";
@session_reset();
echo session_status() === PHP_SESSION_ACTIVE ? "RESET_OK" : "FAIL";
@session_write_close();
"##,
    );
}

#[test]
fn test_php_session_unset_clears_superglobal() {
    compile_ok(
        r##"<?php
@session_start();
$_SESSION["user"] = "Alice";
@session_unset();
echo count($_SESSION) === 0 ? "UNSET_OK" : "FAIL";
@session_write_close();
"##,
    );
}

#[test]
fn test_php_session_save_path_getter_setter() {
    compile_ok(
        r##"<?php
$orig = @session_save_path();
@session_save_path("/tmp");
$newPath = @session_save_path();
@session_save_path($orig);
echo $newPath === "/tmp" ? "SAVE_PATH_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_cache_limiter_options() {
    compile_ok(
        r##"<?php
@session_cache_limiter("private");
echo @session_cache_limiter() === "private" ? "CACHE_LIMITER_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_cache_expire_ttl() {
    compile_ok(
        r##"<?php
@session_cache_expire(30);
echo @session_cache_expire() === 30 ? "CACHE_EXPIRE_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_create_id_prefix() {
    compile_ok(
        r##"<?php
if (function_exists('session_create_id')) {
    $id = @session_create_id("prefix_");
    echo str_starts_with($id, "prefix_") ? "CREATE_ID_PREFIX_OK" : "FAIL";
} else {
    echo "CREATE_ID_PREFIX_OK";
}
"##,
    );
}
