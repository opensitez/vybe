use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Filter Validation & Sanitization — filter_var, filter_var_array, FILTER_VALIDATE_EMAIL, FILTER_VALIDATE_URL, FILTER_VALIDATE_INT
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_filter_var_email_validation() {
    let out = run_prints(
        r#"<?php
$valid = "user@example.com";
$invalid = "invalid-email-str";

echo filter_var($valid, FILTER_VALIDATE_EMAIL) ? "VALID" : "INVALID";
echo " | ";
echo filter_var($invalid, FILTER_VALIDATE_EMAIL) ? "VALID" : "INVALID";
"#,
    );
    assert_eq!(out, vec!["VALID | INVALID"]);
}

#[test]
fn test_php_filter_var_url_validation() {
    let out = run_prints(
        r#"<?php
$url = "https://laravel.com/docs/10.x";
echo filter_var($url, FILTER_VALIDATE_URL) ? "VALID_URL" : "INVALID_URL";
"#,
    );
    assert_eq!(out, vec!["VALID_URL"]);
}

#[test]
fn test_php_filter_var_int_range_options() {
    let out = run_prints(
        r#"<?php
$val = 42;
$options = ["options" => ["min_range" => 1, "max_range" => 50]];
echo filter_var($val, FILTER_VALIDATE_INT, $options) !== false ? "IN_RANGE" : "OUT_OF_RANGE";
"#,
    );
    assert_eq!(out, vec!["IN_RANGE"]);
}

#[test]
fn test_php_filter_var_boolean_parsing() {
    let out = run_prints(
        r#"<?php
$trueVals = ["true", "1", "yes", "on"];
$results = [];
foreach ($trueVals as $v) {
    $results[] = filter_var($v, FILTER_VALIDATE_BOOLEAN) ? "1" : "0";
}
echo implode("-", $results);
"#,
    );
    assert_eq!(out, vec!["1-1-1-1"]);
}

#[test]
fn test_php_filter_var_array_bulk_validation() {
    compile_ok(
        r#"<?php
$data = [
    "email" => "test@domain.org",
    "age" => "25",
    "website" => "http://site.com",
];

$definition = [
    "email" => FILTER_VALIDATE_EMAIL,
    "age" => ["filter" => FILTER_VALIDATE_INT, "options" => ["min_range" => 18]],
    "website" => FILTER_VALIDATE_URL,
];

$validated = filter_var_array($data, $definition);
print_r($validated);
"#,
    );
}

#[test]
fn test_php_filter_var_ip_address_v4_v6_flags() {
    compile_ok(
        r#"<?php
$ipv4 = "192.168.1.1";
$ipv6 = "2001:0db8:85a3:0000:0000:8a2e:0370:7334";

echo filter_var($ipv4, FILTER_VALIDATE_IP, FILTER_FLAG_IPV4) ? "V4_OK" : "V4_FAIL";
echo filter_var($ipv6, FILTER_VALIDATE_IP, FILTER_FLAG_IPV6) ? "V6_OK" : "V6_FAIL";
"#,
    );
}

#[test]
fn test_php_filter_var_domain_name_validation() {
    compile_ok(
        r#"<?php
$domain = "subdomain.example.co.uk";
echo filter_var($domain, FILTER_VALIDATE_DOMAIN) ? "DOMAIN_OK" : "DOMAIN_FAIL";
"#,
    );
}

#[test]
fn test_php_filter_var_float_fraction_validation() {
    compile_ok(
        r#"<?php
$floatStr = "3.14159";
$res = filter_var($floatStr, FILTER_VALIDATE_FLOAT);
echo $res !== false ? "FLOAT_OK" : "FLOAT_FAIL";
"#,
    );
}

#[test]
fn test_php_filter_var_null_on_failure_flag() {
    compile_ok(
        r#"<?php
$res = filter_var("not_a_number", FILTER_VALIDATE_INT, FILTER_NULL_ON_FAILURE);
echo is_null($res) ? "NULL_ON_FAIL" : "OTHER";
"#,
    );
}

#[test]
fn test_php_filter_input_get_post_globals() {
    compile_ok(
        r#"<?php
$_GET["id"] = "100";
$id = filter_input(INPUT_GET, "id", FILTER_VALIDATE_INT);
echo $id;
"#,
    );
}
