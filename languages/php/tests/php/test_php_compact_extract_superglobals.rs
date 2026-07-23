use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Compact, Extract & Superglobals — compact, extract, EXTR_OVERWRITE, $_GET, $_POST, $_SERVER, $_COOKIE, $_FILES, $_SESSION, $_ENV
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_compact_create_array_from_variables() {
    let out = run_prints(
        r#"<?php
$city = "Paris";
$country = "France";
$population = 2161000;

$data = compact("city", "country", "population");
echo "{$data['city']} in {$data['country']} pop={$data['population']}";
"#,
    );
    assert_eq!(out, vec!["Paris in France pop=2161000"]);
}

#[test]
fn test_php_extract_import_array_into_symbol_table() {
    let out = run_prints(
        r#"<?php
$user = ["username" => "alice", "role" => "admin"];
extract($user);

echo "$username is $role";
"#,
    );
    assert_eq!(out, vec!["alice is admin"]);
}

#[test]
fn test_php_extract_prefix_all_collision_avoidance() {
    let out = run_prints(
        r#"<?php
$name = "Original Name";
$params = ["name" => "New Name", "id" => 100];

extract($params, EXTR_PREFIX_ALL, "req");
echo "$name | $req_name | $req_id";
"#,
    );
    assert_eq!(out, vec!["Original Name | New Name | 100"]);
}

#[test]
fn test_php_compact_nested_array_argument() {
    let out = run_prints(
        r#"<?php
$a = 1;
$b = 2;
$c = 3;
$keys = ["a", ["b", "c"]];
$result = compact($keys);
echo implode(",", array_keys($result));
"#,
    );
    assert_eq!(out, vec!["a,b,c"]);
}

#[test]
fn test_php_superglobals_server_get_post_access() {
    let out = run_prints(
        r#"<?php
$_SERVER["REQUEST_METHOD"] = "POST";
$_GET["page"] = "2";
$_POST["token"] = "abc123token";

echo $_SERVER["REQUEST_METHOD"] . " page=" . $_GET["page"] . " token=" . $_POST["token"];
"#,
    );
    assert_eq!(out, vec!["POST page=2 token=abc123token"]);
}

#[test]
fn test_php_extract_skip_existing_variables() {
    compile_ok(
        r#"<?php
$status = "protected";
$input = ["status" => "overwritten", "new_key" => "value"];

extract($input, EXTR_SKIP);
echo "$status | $new_key";
"#,
    );
}

#[test]
fn test_php_superglobals_env_cookie_files_access() {
    compile_ok(
        r#"<?php
$_ENV["APP_KEY"] = "base64:secret";
$_COOKIE["session"] = "cookie_val";
$_FILES["upload"] = ["name" => "photo.jpg", "size" => 1024];

echo $_ENV["APP_KEY"] . " " . $_COOKIE["session"] . " " . $_FILES["upload"]["name"];
"#,
    );
}

#[test]
fn test_php_get_defined_vars_scope_inspection() {
    compile_ok(
        r#"<?php
$x = 10;
$y = "hello";
$vars = get_defined_vars();
echo isset($vars["x"]) && isset($vars["y"]) ? "VARS_FOUND" : "FAIL";
"#,
    );
}

#[test]
fn test_php_globals_array_superglobal_access() {
    compile_ok(
        r#"<?php
$globalVar = "Global Scope";
function testGlobal() {
    echo $GLOBALS["globalVar"];
}
testGlobal();
"#,
    );
}

#[test]
fn test_php_extract_if_exists_import() {
    compile_ok(
        r#"<?php
$existing = "initial";
$input = ["existing" => "updated", "non_existing" => "ignored"];

extract($input, EXTR_IF_EXISTS);
echo "$existing " . (isset($non_existing) ? "YES" : "NO");
"#,
    );
}
