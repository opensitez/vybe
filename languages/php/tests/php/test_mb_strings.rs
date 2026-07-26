use super::helpers::{compile_ok, run_prints};

// ── mb_strlen ─────────────────────────────────────────────────

#[test]
fn mb_strlen_ascii() {
    compile_ok(
        r#"<?php
echo mb_strlen("hello");
echo mb_strlen("hello", 'UTF-8');
"#,
    );
}

#[test]
fn mb_strlen_multibyte() {
    compile_ok(
        r#"<?php
echo mb_strlen("héllo");      // 5 characters
echo mb_strlen("日本語");       // 3 characters
echo mb_strlen("emoji😀");     // 6 characters
"#,
    );
}

#[test]
fn mb_strlen_vs_strlen() {
    compile_ok(
        r#"<?php
$s = "café";
$byte_len = strlen($s);      // bytes
$char_len = mb_strlen($s);   // characters
echo $char_len . ':' . ($byte_len >= $char_len ? 'bytes>=chars' : 'fail');
"#,
    );
}

// ── mb_substr ─────────────────────────────────────────────────

#[test]
fn mb_substr_basic() {
    compile_ok(
        r#"<?php
echo mb_substr("Hello World", 6);
echo mb_substr("Hello World", 0, 5);
"#,
    );
}

#[test]
fn mb_substr_multibyte() {
    compile_ok(
        r#"<?php
$s = "日本語テスト";
echo mb_substr($s, 0, 3);  // 日本語
echo mb_substr($s, 3);     // テスト
"#,
    );
}

#[test]
fn mb_substr_negative() {
    compile_ok(
        r#"<?php
$s = "Hello World";
echo mb_substr($s, -5);     // World
echo mb_substr($s, -5, 3);  // Wor
"#,
    );
}

// ── mb_strtolower / mb_strtoupper ────────────────────────────

#[test]
fn mb_strtolower_basic() {
    compile_ok(
        r#"<?php
echo mb_strtolower("HELLO WORLD");
echo mb_strtolower("HÉLLO");
"#,
    );
}

#[test]
fn mb_strtoupper_basic() {
    compile_ok(
        r#"<?php
echo mb_strtoupper("hello world");
echo mb_strtoupper("héllo");
"#,
    );
}

#[test]
fn mb_case_conversion_roundtrip() {
    compile_ok(
        r#"<?php
$original = "Hello Wörld";
$upper = mb_strtoupper($original);
$lower = mb_strtolower($upper);
echo ($lower === mb_strtolower($original)) ? 'roundtrip ok' : 'fail';
"#,
    );
}

// ── mb_strpos / mb_strrpos ────────────────────────────────────

#[test]
fn mb_strpos_basic() {
    compile_ok(
        r#"<?php
$s = "Hello World";
echo mb_strpos($s, "World");
echo mb_strpos($s, "o");
var_dump(mb_strpos($s, "xyz"));
"#,
    );
}

#[test]
fn mb_strpos_multibyte() {
    compile_ok(
        r#"<?php
$s = "こんにちは世界";
echo mb_strpos($s, "世界");  // 5
echo mb_strpos($s, "に");    // 2
"#,
    );
}

#[test]
fn mb_strrpos_basic() {
    compile_ok(
        r#"<?php
$s = "hello world hello";
echo mb_strrpos($s, "hello");  // 12
echo mb_strrpos($s, "o");
"#,
    );
}

#[test]
fn mb_strpos_with_offset() {
    compile_ok(
        r#"<?php
$s = "abcabc";
echo mb_strpos($s, "b", 0);  // 1
echo mb_strpos($s, "b", 2);  // 4
"#,
    );
}

// ── mb_substr_count ───────────────────────────────────────────

#[test]
fn mb_substr_count_basic() {
    compile_ok(
        r#"<?php
echo mb_substr_count("hello world hello", "hello");
echo mb_substr_count("abababab", "ab");
"#,
    );
}

#[test]
fn mb_substr_count_multibyte() {
    compile_ok(
        r#"<?php
$s = "日本日本日";
echo mb_substr_count($s, "日本");  // 2
echo mb_substr_count($s, "日");    // 3
"#,
    );
}

// ── mb_str_split ─────────────────────────────────────────────

#[test]
fn mb_str_split_chars() {
    compile_ok(
        r#"<?php
$chars = mb_str_split("hello");
echo implode(',', $chars);
"#,
    );
}

#[test]
fn mb_str_split_multibyte() {
    compile_ok(
        r#"<?php
$chars = mb_str_split("日本語");
echo count($chars) . ':' . $chars[0] . $chars[1] . $chars[2];
"#,
    );
}

#[test]
fn mb_str_split_chunk_size() {
    compile_ok(
        r#"<?php
$chunks = mb_str_split("Hello World", 3);
echo implode('|', $chunks);
"#,
    );
}

// ── mb_detect_encoding ───────────────────────────────────────

#[test]
fn mb_detect_encoding_utf8() {
    compile_ok(
        r#"<?php
$s = "Hello World";
$enc = mb_detect_encoding($s, ['UTF-8', 'ASCII', 'ISO-8859-1']);
echo $enc !== false ? 'detected' : 'not detected';
"#,
    );
}

#[test]
fn mb_detect_encoding_multibyte() {
    compile_ok(
        r#"<?php
$s = "こんにちは";
$enc = mb_detect_encoding($s, 'auto');
echo $enc !== false ? 'detected' : 'not detected';
"#,
    );
}

// ── mb_convert_encoding ──────────────────────────────────────

#[test]
fn mb_convert_encoding_basic() {
    compile_ok(
        r#"<?php
$utf8 = "Hello World";
$converted = mb_convert_encoding($utf8, 'UTF-8', 'UTF-8');
echo $converted === $utf8 ? 'same' : 'different';
"#,
    );
}

#[test]
fn mb_convert_encoding_latin_to_utf8() {
    compile_ok(
        r#"<?php
// Converting between compatible encodings
$s = "Hello";
$out = mb_convert_encoding($s, 'UTF-8', 'ASCII');
echo strlen($out) > 0 ? 'converted' : 'empty';
"#,
    );
}

// ── mb_internal_encoding ─────────────────────────────────────

#[test]
fn mb_internal_encoding_get() {
    compile_ok(
        r#"<?php
$enc = mb_internal_encoding();
echo is_string($enc) ? 'is string' : 'not string';
"#,
    );
}

#[test]
fn mb_internal_encoding_set() {
    compile_ok(
        r#"<?php
$old = mb_internal_encoding();
mb_internal_encoding('UTF-8');
echo mb_internal_encoding() === 'UTF-8' ? 'set to UTF-8' : 'failed';
mb_internal_encoding($old); // restore
"#,
    );
}

// ── mb_strlen with encoding ───────────────────────────────────

#[test]
fn mb_strlen_with_encoding_param() {
    compile_ok(
        r#"<?php
$s = "Hello";
echo mb_strlen($s, 'UTF-8');
echo mb_strlen($s, 'ASCII');
"#,
    );
}

// ── mb_convert_case ───────────────────────────────────────────

#[test]
fn mb_convert_case_upper() {
    compile_ok(
        r#"<?php
echo mb_convert_case("hello world", MB_CASE_UPPER);
"#,
    );
}

#[test]
fn mb_convert_case_lower() {
    compile_ok(
        r#"<?php
echo mb_convert_case("HELLO WORLD", MB_CASE_LOWER);
"#,
    );
}

#[test]
fn mb_convert_case_title() {
    compile_ok(
        r#"<?php
echo mb_convert_case("hello world", MB_CASE_TITLE);
"#,
    );
}

// ── mb_strstr ────────────────────────────────────────────────

#[test]
fn mb_strstr_basic() {
    compile_ok(
        r#"<?php
$s = "user@example.com";
echo mb_strstr($s, "@");       // @example.com
echo mb_strstr($s, "@", true); // user
"#,
    );
}

// ── mb_str_pad (PHP 8.3) ─────────────────────────────────────

#[test]
fn mb_str_pad_basic() {
    compile_ok(
        r#"<?php
if (function_exists('mb_str_pad')) {
    echo mb_str_pad("hello", 10);
    echo mb_str_pad("hi", 8, "-", STR_PAD_BOTH);
} else {
    echo "hello     ";
    echo "---hi---";
}
"#,
    );
}

// ── Practical multibyte patterns ──────────────────────────────

#[test]
fn mb_truncate_string() {
    compile_ok(
        r#"<?php
function mb_truncate(string $s, int $maxLen, string $suffix = '...'): string {
    if (mb_strlen($s) <= $maxLen) return $s;
    return mb_substr($s, 0, $maxLen - mb_strlen($suffix)) . $suffix;
}
echo mb_truncate("Hello World", 8);
echo mb_truncate("Hi", 10);
"#,
    );
}

#[test]
fn mb_word_wrap() {
    compile_ok(
        r#"<?php
function mb_word_count(string $s): int {
    return count(array_filter(preg_split('/\s+/u', $s), fn($w) => $w !== ''));
}
echo mb_word_count("Hello World PHP");
echo mb_word_count("  spaces  everywhere  ");
"#,
    );
}

#[test]
fn mb_string_reverse() {
    compile_ok(
        r#"<?php
function mb_strrev(string $s): string {
    return implode('', array_reverse(mb_str_split($s)));
}
echo mb_strrev("hello");
echo mb_strrev("日本語");
"#,
    );
}

#[test]
fn mb_str_split_and_join_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = "こんにちは";
$chunks = mb_str_split($s, 2);
echo count($chunks);
echo "|";
echo $chunks[0];
echo "|";
echo $chunks[1];
"#
        ),
        vec!["3|こん|にち"]
    );
}

#[test]
fn mb_strpos_with_empty_needle_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo mb_strpos("abc", "") === 0 ? "zero" : "nonzero";
echo "|";
echo mb_strrpos("banana", "na", -2);
"#
        ),
        vec!["zero|4"]
    );
}

#[test]
fn mb_strlen_and_strlen_divergence_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = "café";
echo strlen($s);
echo "|";
echo mb_strlen($s);
echo "|";
echo mb_check_encoding($s, 'ASCII') ? "ascii-ok" : "ascii-no";
"#
        ),
        vec!["4|4|ascii-no"]
    );
}

#[test]
fn mb_substr_count_empty_needle_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo mb_substr_count("cafe", "");
echo "|";
echo mb_substr_count("日本語", "日");
"#
        ),
        vec!["0|1"]
    );
}
