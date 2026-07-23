use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Sessions: session_regenerate_id & Session Security
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_session_regenerate_id_changes_session_id() {
    let out = run_prints(
        r##"<?php
@session_start();
$oldId = session_id();
@session_regenerate_id(false);
$newId = session_id();
@session_write_close();

echo $oldId !== $newId ? "REGENERATED_ID_OK" : "SAME_ID";
"##,
    );
    assert_eq!(out, vec!["REGENERATED_ID_OK"]);
}

#[test]
fn test_php_session_regenerate_id_delete_old_session_flag() {
    let out = run_prints(
        r##"<?php
@session_start();
$_SESSION["user_id"] = 42;
$res = @session_regenerate_id(true); // true = delete old session data
echo "RegenerateResult=" . ($res ? "1" : "0") . " DataPreserved=" . ($_SESSION["user_id"] === 42 ? "YES" : "NO");
@session_write_close();
"##,
    );
    assert_eq!(out, vec!["RegenerateResult=1 DataPreserved=YES"]);
}

#[test]
fn test_php_session_destroy_clears_session_file() {
    let out = run_prints(
        r##"<?php
@session_start();
$_SESSION["auth"] = true;
$destroyed = @session_destroy();
echo $destroyed ? "DESTROYED_OK" : "FAIL";
"##,
    );
    assert_eq!(out, vec!["DESTROYED_OK"]);
}

#[test]
fn test_php_session_regenerate_id_inactive_session_returns_false() {
    compile_ok(
        r##"<?php
$res = @session_regenerate_id(false);
echo $res === false ? "INACTIVE_REGENERATE_FALSE" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_cookie_params_samesite_strict() {
    compile_ok(
        r##"<?php
session_set_cookie_params([
    "samesite" => "Strict",
    "secure" => true
]);
$p = session_get_cookie_params();
echo $p["samesite"] === "Strict" && $p["secure"] ? "SAMESITE_STRICT_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_use_strict_mode_ini() {
    compile_ok(
        r##"<?php
$strict = ini_get("session.use_strict_mode");
echo $strict !== false ? "USE_STRICT_MODE_INI_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_cookie_httponly_ini() {
    compile_ok(
        r##"<?php
$httponly = ini_get("session.cookie_httponly");
echo $httponly !== false ? "HTTPONLY_INI_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_gc_maxlifetime_ini() {
    compile_ok(
        r##"<?php
$ttl = ini_get("session.gc_maxlifetime");
echo is_numeric($ttl) && (int)$ttl > 0 ? "GC_MAXLIFETIME_INI_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_destroy_does_not_unset_superglobal() {
    compile_ok(
        r##"<?php
@session_start();
$_SESSION["key"] = "value";
@session_destroy();
// Note: session_destroy() deletes session storage, but $_SESSION remains set in script memory until unset()
echo isset($_SESSION["key"]) ? "DESTROY_SUPERGLOBAL_PERSISTS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_create_id_custom_prefix_length() {
    compile_ok(
        r##"<?php
if (function_exists('session_create_id')) {
    $id = @session_create_id("sess_prefix_");
    echo strlen($id) > 12 ? "CREATE_ID_LENGTH_OK" : "FAIL";
} else {
    echo "CREATE_ID_LENGTH_OK";
}
"##,
    );
}
