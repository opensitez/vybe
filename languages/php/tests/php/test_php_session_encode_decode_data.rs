use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Sessions: session_encode & session_decode Data Serialization
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_session_encode_serializes_active_session() {
    let out = run_prints(
        r##"<?php
@session_start();
$_SESSION["username"] = "Alice";
$_SESSION["role"] = "admin";

$encoded = @session_encode();
@session_write_close();

echo str_contains($encoded, "username") && str_contains($encoded, "Alice") ? "ENCODED_OK" : "FAIL";
"##,
    );
    assert_eq!(out, vec!["ENCODED_OK"]);
}

#[test]
fn test_php_session_decode_restores_session_variables() {
    let out = run_prints(
        r##"<?php
@session_start();
$_SESSION = [];
$data = 'user|s:3:"Bob";id|i:42;';
@session_decode($data);

echo "User=" . ($_SESSION["user"] ?? "") . " ID=" . ($_SESSION["id"] ?? 0);
@session_write_close();
"##,
    );
    assert_eq!(out, vec!["User=Bob ID=42"]);
}

#[test]
fn test_php_session_encode_decode_roundtrip() {
    compile_ok(
        r##"<?php
@session_start();
$_SESSION["items"] = ["item1", "item2"];
$encoded = @session_encode();

$_SESSION = [];
@session_decode($encoded);

echo count($_SESSION["items"] ?? []) === 2 ? "ROUNDTRIP_OK" : "FAIL";
@session_write_close();
"##,
    );
}

#[test]
fn test_php_session_serialize_handler_ini_setting() {
    compile_ok(
        r##"<?php
$handler = ini_get("session.serialize_handler");
echo is_string($handler) && strlen($handler) > 0 ? "SERIALIZE_HANDLER_INI_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_decode_empty_string_clears_data() {
    compile_ok(
        r##"<?php
@session_start();
$_SESSION["data"] = 123;
@session_decode("");
echo count($_SESSION) === 0 ? "DECODE_EMPTY_CLEARS_OK" : "FAIL";
@session_write_close();
"##,
    );
}

#[test]
fn test_php_session_encode_nested_array_structures() {
    compile_ok(
        r##"<?php
@session_start();
$_SESSION["config"] = ["db" => ["host" => "localhost", "port" => 3306]];
$encoded = @session_encode();
$_SESSION = [];
@session_decode($encoded);
echo $_SESSION["config"]["db"]["port"] === 3306 ? "NESTED_ENCODE_OK" : "FAIL";
@session_write_close();
"##,
    );
}

#[test]
fn test_php_session_encode_boolean_and_null_types() {
    compile_ok(
        r##"<?php
@session_start();
$_SESSION["flag"] = true;
$_SESSION["nil"] = null;
$encoded = @session_encode();
$_SESSION = [];
@session_decode($encoded);
echo $_SESSION["flag"] === true && $_SESSION["nil"] === null ? "BOOL_NULL_DECODE_OK" : "FAIL";
@session_write_close();
"##,
    );
}

#[test]
fn test_php_session_get_cookie_params_structure() {
    compile_ok(
        r##"<?php
$params = session_get_cookie_params();
echo isset($params["lifetime"]) && isset($params["path"]) && isset($params["domain"]) ? "COOKIE_PARAMS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_set_cookie_params_options_array() {
    compile_ok(
        r##"<?php
session_set_cookie_params([
    "lifetime" => 7200,
    "path" => "/app",
    "secure" => true,
    "httponly" => true,
    "samesite" => "Lax"
]);
$p = session_get_cookie_params();
echo $p["lifetime"] === 7200 && $p["path"] === "/app" ? "SET_COOKIE_PARAMS_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_session_abort_discards_changes() {
    compile_ok(
        r##"<?php
@session_start();
$_SESSION["temp"] = "discard_me";
@session_abort();
echo session_status() === PHP_SESSION_NONE ? "ABORT_STATUS_NONE_OK" : "FAIL";
"##,
    );
}
