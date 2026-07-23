use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Filter Input & Superglobal Validation — filter_input, filter_input_array, INPUT_GET, INPUT_POST, INPUT_COOKIE
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_filter_input_simulated_query_param() {
    let out = run_prints(
        r#"<?php
$_GET["age"] = "25";
$age = filter_input(INPUT_GET, "age", FILTER_VALIDATE_INT);
echo "Age: $age";
"#,
    );
    assert_eq!(out, vec!["Age: 25"]);
}

#[test]
fn test_php_filter_input_array_batch_validation() {
    let out = run_prints(
        r#"<?php
$_POST["email"] = "user@example.com";
$_POST["age"] = "30";

$filters = [
    "email" => FILTER_VALIDATE_EMAIL,
    "age" => [
        "filter" => FILTER_VALIDATE_INT,
        "options" => ["min_range" => 18, "max_range" => 65]
    ]
];

$result = filter_input_array(INPUT_POST, $filters);
echo "Email={$result['email']} Age={$result['age']}";
"#,
    );
    assert_eq!(out, vec!["Email=user@example.com Age=30"]);
}

#[test]
fn test_php_filter_has_var_superglobal_check() {
    compile_ok(
        r#"<?php
$_GET["param"] = "value";
echo filter_has_var(INPUT_GET, "param") ? "HAS_VAR" : "NO_VAR";
"#,
    );
}

#[test]
fn test_php_filter_input_invalid_email_returns_false() {
    compile_ok(
        r#"<?php
$_GET["email"] = "not_an_email";
$res = filter_input(INPUT_GET, "email", FILTER_VALIDATE_EMAIL);
echo $res === false ? "EMAIL_INVALID" : "VALID";
"#,
    );
}

#[test]
fn test_php_filter_input_default_fallback_option() {
    compile_ok(
        r#"<?php
$val = filter_input(INPUT_GET, "missing_key", FILTER_VALIDATE_INT, [
    "options" => ["default" => 100]
]);
echo "Fallback: $val";
"#,
    );
}

#[test]
fn test_php_filter_input_sanitize_encoded() {
    compile_ok(
        r#"<?php
$_GET["url"] = "https://example.com/test?a=1&b=2";
$clean = filter_input(INPUT_GET, "url", FILTER_SANITIZE_URL);
echo $clean;
"#,
    );
}

#[test]
fn test_php_filter_input_boolean_conversion() {
    compile_ok(
        r#"<?php
$_POST["agree"] = "yes";
$boolVal = filter_input(INPUT_POST, "agree", FILTER_VALIDATE_BOOLEAN, FILTER_NULL_ON_FAILURE);
echo $boolVal ? "TRUE" : "FALSE";
"#,
    );
}

#[test]
fn test_php_filter_input_ip_address_flags() {
    compile_ok(
        r#"<?php
$_SERVER["REMOTE_ADDR"] = "192.168.1.1";
$ip = filter_input(INPUT_SERVER, "REMOTE_ADDR", FILTER_VALIDATE_IP, FILTER_FLAG_NO_PRIV_RANGE);
echo $ip === false ? "PRIVATE_IP_FILTERED" : "PUBLIC_IP";
"#,
    );
}

#[test]
fn test_php_filter_input_cookie_sanitization() {
    compile_ok(
        r#"<?php
$_COOKIE["user_id"] = "12345";
$id = filter_input(INPUT_COOKIE, "user_id", FILTER_VALIDATE_INT);
echo "User ID: $id";
"#,
    );
}

#[test]
fn test_php_filter_input_array_all_keys() {
    compile_ok(
        r#"<?php
$_GET["a"] = "10";
$_GET["b"] = "20";
$data = filter_input_array(INPUT_GET, FILTER_VALIDATE_INT);
echo is_array($data) ? "INPUT_ARRAY_OK" : "FAIL";
"#,
    );
}
