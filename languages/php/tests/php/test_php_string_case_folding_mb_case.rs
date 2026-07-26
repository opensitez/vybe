use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: String Case Folding & Multibyte Case Conversion — ucfirst, lcfirst, ucwords, mb_convert_case, MB_CASE_TITLE, MB_CASE_FOLD
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_ucfirst_lcfirst_ucwords_transformation() {
    let out = run_prints(
        r#"<?php
$str = "hello world";
echo ucfirst($str) . " | " . ucwords($str) . " | " . lcfirst("HELLO");
"#,
    );
    assert_eq!(out, vec!["Hello world | Hello World | hELLO"]);
}

#[test]
fn test_php_mb_convert_case_title_upper_lower() {
    let out = run_prints(
        r#"<?php
$title = "münchen & köln";
$formatted = mb_convert_case($title, MB_CASE_TITLE, "UTF-8");
echo $formatted;
"#,
    );
    assert_eq!(out, vec!["München & Köln"]);
}

#[test]
fn test_php_mb_convert_case_fold_case_folding() {
    let out = run_prints(
        r#"<?php
$greek = "ΟΔΥΣΣΕΥΣ";
$folded = mb_convert_case($greek, MB_CASE_FOLD, "UTF-8");
echo mb_strlen($folded, "UTF-8") > 0 ? "FOLD_OK" : "FOLD_FAIL";
"#,
    );
    assert_eq!(out, vec!["FOLD_OK"]);
}

#[test]
fn test_php_ucwords_custom_delimiters() {
    let out = run_prints(
        r#"<?php
$str = "hello|world|php";
echo ucwords($str, "|");
"#,
    );
    assert_eq!(out, vec!["Hello|World|Php"]);
}

#[test]
fn test_php_strtolower_strtoupper_ascii() {
    compile_ok(
        r#"<?php
$raw = "Laravel Framework 10.x";
echo strtolower($raw) . " " . strtoupper($raw);
"#,
    );
}

#[test]
fn test_php_mb_convert_case_upper_lower_modes() {
    compile_ok(
        r#"<?php
$text = "éclair";
echo mb_convert_case($text, MB_CASE_UPPER, "UTF-8") . " " . mb_convert_case("ÉCLAIR", MB_CASE_LOWER, "UTF-8");
"#,
    );
}

#[test]
fn test_php_lcfirst_multibyte_fallback() {
    compile_ok(
        r#"<?php
echo lcfirst("World");
"#,
    );
}

#[test]
fn test_php_mb_case_title_lower_mapping() {
    compile_ok(
        r#"<?php
$text = "THE QUICK BROWN FOX";
echo mb_convert_case($text, MB_CASE_TITLE, "UTF-8");
"#,
    );
}

#[test]
fn test_php_string_case_sensitivity_helpers() {
    compile_ok(
        r#"<?php
$a = "test";
$b = "TEST";
echo strcasecmp($a, $b) === 0 ? "EQUAL_NOCASE" : "NOT_EQUAL";
"#,
    );
}

#[test]
fn test_php_strncasecmp_length_limited() {
    compile_ok(
        r#"<?php
echo strncasecmp("Hello World", "hello php", 5) === 0 ? "MATCH_5" : "NO_MATCH";
"#,
    );
}

#[test]
fn test_php_strtolower_preserves_spaces_and_punctuation() {
    let out = run_prints(
        r#"<?php
echo strtolower("  Hello, World!  ");
echo "|";
echo strtoupper("café");
"#,
    );
    assert_eq!(out, vec!["  hello, world!  |CAFÉ"]);
}

#[test]
fn test_php_mb_convert_case_multibyte_modes_runtime() {
    let out = run_prints(
        r#"<?php
echo mb_convert_case("straße", MB_CASE_UPPER, "UTF-8");
echo "|";
echo mb_convert_case("İ", MB_CASE_LOWER, "UTF-8");
"#,
    );
    assert_eq!(out, vec!["STRASSE|i̇"]);
}

#[test]
fn test_php_ucfirst_empty_and_singleton() {
    let out = run_prints(
        r#"<?php
echo ucfirst("") === "" ? "empty" : "no";
echo "|";
echo ucfirst("δ");
echo "|";
echo ucfirst("ß");
"#,
    );
    assert_eq!(out, vec!["empty|δ|ß"]);
}
