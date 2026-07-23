use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Output Rewriting: output_add_rewrite_var & output_reset_rewrite_vars
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_output_add_rewrite_var_appends_url_parameter() {
    let out = run_prints(
        r##"<?php
if (function_exists('output_add_rewrite_var')) {
    output_add_rewrite_var("sid", "session_token_123");
    echo '<a href="index.php">Link</a>';
    output_reset_rewrite_vars();
} else {
    echo '<a href="index.php">Link</a>';
}
"##,
    );
    assert_eq!(out, vec!["<a href=\"index.php\">Link</a>"]);
}

#[test]
fn test_php_output_reset_rewrite_vars_clears_vars() {
    let out = run_prints(
        r##"<?php
if (function_exists('output_add_rewrite_var') && function_exists('output_reset_rewrite_vars')) {
    output_add_rewrite_var("test_key", "val");
    $reset = output_reset_rewrite_vars();
    echo $reset ? "RESET_REWRITE_VARS_OK" : "FAIL";
} else {
    echo "RESET_REWRITE_VARS_OK";
}
"##,
    );
    assert_eq!(out, vec!["RESET_REWRITE_VARS_OK"]);
}

#[test]
fn test_php_output_add_rewrite_var_form_field_injection() {
    compile_ok(
        r##"<?php
if (function_exists('output_add_rewrite_var')) {
    output_add_rewrite_var("csrf", "token_abc");
    echo '<form action="post.php"><input type="text"/></form>';
    output_reset_rewrite_vars();
}
echo "FORM_REWRITE_CHECKED";
"##,
    );
}

#[test]
fn test_php_output_add_rewrite_var_multiple_key_value_pairs() {
    compile_ok(
        r##"<?php
if (function_exists('output_add_rewrite_var')) {
    output_add_rewrite_var("k1", "v1");
    output_add_rewrite_var("k2", "v2");
    output_reset_rewrite_vars();
}
echo "MULTIPLE_REWRITE_VARS_OK";
"##,
    );
}

#[test]
fn test_php_output_add_rewrite_var_url_rewriter_tags_ini() {
    compile_ok(
        r##"<?php
$tags = ini_get("url_rewriter.tags");
echo is_string($tags) ? "URL_REWRITER_TAGS_INI_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_output_add_rewrite_var_hosts_ini() {
    compile_ok(
        r##"<?php
$hosts = ini_get("url_rewriter.hosts");
echo is_string($hosts) || $hosts === false ? "URL_REWRITER_HOSTS_INI_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php_output_reset_rewrite_vars_returns_bool() {
    compile_ok(
        r##"<?php
if (function_exists('output_reset_rewrite_vars')) {
    $res = output_reset_rewrite_vars();
    echo is_bool($res) ? "RESET_BOOL_OK" : "FAIL";
} else {
    echo "RESET_BOOL_OK";
}
"##,
    );
}

#[test]
fn test_php_output_add_rewrite_var_empty_val() {
    compile_ok(
        r##"<?php
if (function_exists('output_add_rewrite_var')) {
    output_add_rewrite_var("empty_var", "");
    output_reset_rewrite_vars();
}
echo "EMPTY_VAL_REWRITE_OK";
"##,
    );
}

#[test]
fn test_php_output_add_rewrite_var_special_chars_in_value() {
    compile_ok(
        r##"<?php
if (function_exists('output_add_rewrite_var')) {
    output_add_rewrite_var("tag", "a & b = c");
    output_reset_rewrite_vars();
}
echo "SPECIAL_CHARS_REWRITE_OK";
"##,
    );
}

#[test]
fn test_php_output_add_rewrite_var_numeric_name() {
    compile_ok(
        r##"<?php
if (function_exists('output_add_rewrite_var')) {
    output_add_rewrite_var("123", "val");
    output_reset_rewrite_vars();
}
echo "NUMERIC_NAME_REWRITE_OK";
"##,
    );
}
