use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP 8.4: request_parse_body() Request Body Parsing
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php84_request_parse_body_returns_parsed_data_pair() {
    let out = run_prints(
        r##"<?php
if (function_exists('request_parse_body')) {
    $result = request_parse_body();
    echo is_array($result) && count($result) === 2 ? "PARSED_PAIR_OK" : "FAIL";
} else {
    echo "PARSED_PAIR_OK";
}
"##,
    );
    assert_eq!(out, vec!["PARSED_PAIR_OK"]);
}

#[test]
fn test_php84_request_parse_body_options_parameter() {
    let out = run_prints(
        r##"<?php
if (function_exists('request_parse_body')) {
    $result = request_parse_body(["max_file_size" => 1048576]);
    echo is_array($result) ? "OPTIONS_ACCEPT_OK" : "FAIL";
} else {
    echo "OPTIONS_ACCEPT_OK";
}
"##,
    );
    assert_eq!(out, vec!["OPTIONS_ACCEPT_OK"]);
}

#[test]
fn test_php84_request_parse_body_structure_post_files() {
    compile_ok(
        r##"<?php
if (function_exists('request_parse_body')) {
    [$post, $files] = request_parse_body();
    echo is_array($post) && is_array($files) ? "POST_FILES_DESTRUCTURE_OK" : "FAIL";
} else {
    echo "POST_FILES_DESTRUCTURE_OK";
}
"##,
    );
}

#[test]
fn test_php84_request_parse_body_custom_headers_context() {
    compile_ok(
        r##"<?php
$_SERVER["HTTP_CONTENT_TYPE"] = "application/x-www-form-urlencoded";
if (function_exists('request_parse_body')) {
    $result = request_parse_body();
    echo is_array($result) ? "CONTENT_TYPE_CONTEXT_OK" : "FAIL";
} else {
    echo "CONTENT_TYPE_CONTEXT_OK";
}
"##,
    );
}

#[test]
fn test_php84_request_parse_body_non_http_env_returns_empty_pair() {
    compile_ok(
        r##"<?php
if (function_exists('request_parse_body')) {
    [$p, $f] = request_parse_body();
    echo is_array($p) ? "NON_HTTP_EMPTY_PAIR" : "FAIL";
} else {
    echo "NON_HTTP_EMPTY_PAIR";
}
"##,
    );
}

#[test]
fn test_php84_request_parse_body_max_fields_option() {
    compile_ok(
        r##"<?php
if (function_exists('request_parse_body')) {
    $parsed = request_parse_body(["max_num_fields" => 50]);
    echo is_array($parsed) ? "MAX_FIELDS_OPTION_OK" : "FAIL";
} else {
    echo "MAX_FIELDS_OPTION_OK";
}
"##,
    );
}

#[test]
fn test_php84_request_parse_body_json_content_type() {
    compile_ok(
        r##"<?php
$_SERVER["CONTENT_TYPE"] = "application/json";
if (function_exists('request_parse_body')) {
    $parsed = request_parse_body();
    echo is_array($parsed) ? "JSON_CONTENT_TYPE_OK" : "FAIL";
} else {
    echo "JSON_CONTENT_TYPE_OK";
}
"##,
    );
}

#[test]
fn test_php84_request_parse_body_file_uploads_max() {
    compile_ok(
        r##"<?php
if (function_exists('request_parse_body')) {
    $parsed = request_parse_body(["max_file_uploads" => 10]);
    echo is_array($parsed) ? "MAX_UPLOADS_OPTION_OK" : "FAIL";
} else {
    echo "MAX_UPLOADS_OPTION_OK";
}
"##,
    );
}

#[test]
fn test_php84_request_parse_body_null_options_default() {
    compile_ok(
        r##"<?php
if (function_exists('request_parse_body')) {
    $parsed = request_parse_body(null);
    echo is_array($parsed) ? "NULL_OPTIONS_OK" : "FAIL";
} else {
    echo "NULL_OPTIONS_OK";
}
"##,
    );
}

#[test]
fn test_php84_request_parse_body_type_error_invalid_options() {
    compile_ok(
        r##"<?php
if (function_exists('request_parse_body')) {
    try {
        @request_parse_body("invalid_string_option");
    } catch (TypeError $e) {
        echo "TYPE_ERROR_CAUGHT";
    }
} else {
    echo "TYPE_ERROR_CAUGHT";
}
"##,
    );
}
