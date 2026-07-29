use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: String Searching & Substring Positions — strchr, strrchr, strstr, stristr, strrpos, strripos, substr_count, substr_replace
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_strstr_before_and_after_needle() {
    let out = run_prints(
        r#"<?php
$email = "user@example.com";
$domain = strstr($email, "@");
$user = strstr($email, "@", before_needle: true);
echo "$user from $domain";
"#,
    );
    assert_eq!(out, vec!["user from @example.com"]);
}

#[test]
fn test_php_strrchr_last_occurrence_extension() {
    let out = run_prints(
        r#"<?php
$path = "/var/www/html/archive.tar.gz";
$ext = strrchr($path, ".");
echo $ext;
"#,
    );
    assert_eq!(out, vec![".gz"]);
}

#[test]
fn test_php_strrpos_case_sensitive_reverse_search() {
    let out = run_prints(
        r#"<?php
$text = "The quick brown fox jumps over the lazy dog";
$pos = strrpos($text, "the");
echo "Last 'the' at offset: $pos";
"#,
    );
    assert_eq!(out, vec!["Last 'the' at offset: 31"]);
}

#[test]
fn test_php_substr_replace_insertion_and_replacement() {
    let out = run_prints(
        r#"<?php
$var = "ABCDEFGH:/MNODOP/";
echo substr_replace($var, "bob", 3, 4) . " | " . substr_replace($var, "INSERT_", 0, 0);
"#,
    );
    assert_eq!(out, vec!["ABCbobH:/MNODOP/ | INSERT_ABCDEFGH:/MNODOP/"]);
}

#[test]
fn test_php_substr_count_needle_occurrences() {
    let out = run_prints(
        r#"<?php
$text = "This is a test string for testing substr_count";
echo substr_count($text, "is");
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn test_php_stristr_case_insensitive_search() {
    compile_ok(
        r#"<?php
$email = "USER@DOMAIN.COM";
echo stristr($email, "domain.com") ? "MATCH" : "NO_MATCH";
"#,
    );
}

#[test]
fn test_php_strripos_case_insensitive_last_position() {
    compile_ok(
        r#"<?php
$haystack = "Abc ABC abc";
$pos = strripos($haystack, "abc");
echo "Pos=$pos";
"#,
    );
}

#[test]
fn test_php_substr_compare_case_sensitivity_offset() {
    compile_ok(
        r#"<?php
$main = "abcde";
$sub = "BC";
echo substr_compare($main, $sub, 1, 2, case_insensitivity: true) === 0 ? "EQUAL" : "NOT_EQUAL";
"#,
    );
}

#[test]
fn test_php_strpbrk_character_set_search() {
    compile_ok(
        r#"<?php
$text = "This is a Simple text.";
echo strpbrk($text, "mi");
"#,
    );
}

#[test]
fn test_php_substr_replace_array_replacements() {
    compile_ok(
        r#"<?php
$input = ["A: AAA", "B: BBB", "C: CCC"];
$replaced = substr_replace($input, "BBB", 3, 3);
echo implode(",", $replaced);
"#,
    );
}

#[test]
fn test_php_strstr_false_when_needle_missing() {
    let out = run_prints(
        r#"<?php
echo strstr("hello", "z") === false ? "missing" : "found";
echo "|";
echo strstr("hello", "", true) === false ? "empty-before-false" : "empty-before-ok";
"#,
    );
    assert_eq!(out, vec!["missing|empty-before-ok"]);
}

#[test]
fn test_php_strrpos_missing_and_before_offset_behaviour() {
    let out = run_prints(
        r#"<?php
echo var_export(strrpos("banana", "zz"), true);
echo "|";
echo strrpos("ababa", "ba", -3);
"#,
    );
    assert_eq!(out, vec!["false|1"]);
}

#[test]
fn test_php_strripos_with_offset_and_case_miss() {
    let out = run_prints(
        r#"<?php
echo var_export(strripos("Alpha", "Z"), true);
echo "|";
echo strripos("fooABCabc", "ABC", 0);
"#,
    );
    assert_eq!(out, vec!["false|6"]);
}

#[test]
fn test_php_strpbrk_no_match_returns_false() {
    let out = run_prints(
        r#"<?php
echo var_export(strpbrk("hello", "0123"), true);
echo "|";
echo strpbrk("hello", "xol");
"#,
    );
    assert_eq!(out, vec!["false|llo"]);
}

#[test]
fn test_php_substr_compare_boundary_and_negative_offset() {
    let out = run_prints(
        r#"<?php
echo substr_compare("abcdef", "ef", -2);
echo "|";
echo substr_compare("abcdef", "bcd", 1, 3, false);
"#,
    );
    assert_eq!(out, vec!["0|0"]);
}
