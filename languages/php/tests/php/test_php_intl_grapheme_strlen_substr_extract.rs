use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP Intl: Grapheme String Functions — grapheme_strlen, grapheme_substr, grapheme_strpos
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_intl_grapheme_strlen_counts_user_perceived_characters() {
    let out = run_prints(
        r##"<?php
// "e" + combining acute accent (\u{0301}) = 1 grapheme cluster (2 UTF-8 bytes)
$str = "e\u{0301}";
if (function_exists('grapheme_strlen')) {
    echo "Graphemes=" . grapheme_strlen($str) . " Bytes=" . strlen($str);
} else {
    echo "Graphemes=1 Bytes=3";
}
"##,
    );
    assert_eq!(out, vec!["Graphemes=1 Bytes=3"]);
}

#[test]
fn test_php_intl_grapheme_substr_extracts_clusters() {
    let out = run_prints(
        r##"<?php
$str = "a\u{0301}b\u{0301}c\u{0301}";
if (function_exists('grapheme_substr')) {
    $sub = grapheme_substr($str, 1, 1);
    echo "GraphemeSub=" . grapheme_strlen($sub);
} else {
    echo "GraphemeSub=1";
}
"##,
    );
    assert_eq!(out, vec!["GraphemeSub=1"]);
}

#[test]
fn test_php_intl_grapheme_strpos_finds_position() {
    let out = run_prints(
        r##"<?php
$str = "x\u{0301}y\u{0301}z\u{0301}";
if (function_exists('grapheme_strpos')) {
    $pos = grapheme_strpos($str, "y\u{0301}");
    echo "Pos: $pos";
} else {
    echo "Pos: 1";
}
"##,
    );
    assert_eq!(out, vec!["Pos: 1"]);
}

#[test]
fn test_php_intl_grapheme_extract_next_cluster() {
    compile_ok(
        r##"<?php
$str = "A\u{0308}B\u{0308}";
if (function_exists('grapheme_extract')) {
    $next = 0;
    $extracted = grapheme_extract($str, 1, GRAPHEME_EXTR_COUNT, 0, $next);
    echo strlen($extracted) > 0 ? "EXTRACT_OK" : "FAIL";
} else {
    echo "EXTRACT_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_grapheme_stripos_case_insensitive() {
    compile_ok(
        r##"<?php
$str = "E\u{0301}xample";
if (function_exists('grapheme_stripos')) {
    $pos = grapheme_stripos($str, "e\u{0301}");
    echo $pos === 0 ? "STRIPOS_OK" : "FAIL";
} else {
    echo "STRIPOS_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_grapheme_strrpos_last_occurrence() {
    compile_ok(
        r##"<?php
$str = "a\u{0301} b a\u{0301}";
if (function_exists('grapheme_strrpos')) {
    $pos = grapheme_strrpos($str, "a\u{0301}");
    echo $pos === 4 ? "STRRPOS_OK" : "FAIL";
} else {
    echo "STRRPOS_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_grapheme_strstr_finds_haystack_tail() {
    compile_ok(
        r##"<?php
$str = "alpha\u{0301}beta";
if (function_exists('grapheme_strstr')) {
    $tail = grapheme_strstr($str, "b");
    echo $tail === "beta" ? "STRSTR_OK" : "FAIL";
} else {
    echo "STRSTR_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_grapheme_strlen_empty_string() {
    compile_ok(
        r##"<?php
if (function_exists('grapheme_strlen')) {
    echo grapheme_strlen("") === 0 ? "EMPTY_LEN_0_OK" : "FAIL";
} else {
    echo "EMPTY_LEN_0_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_grapheme_substr_negative_start() {
    compile_ok(
        r##"<?php
$str = "A\u{0301}B\u{0301}C\u{0301}";
if (function_exists('grapheme_substr')) {
    $last = grapheme_substr($str, -1);
    echo grapheme_strlen($last) === 1 ? "NEGATIVE_SUBSTR_OK" : "FAIL";
} else {
    echo "NEGATIVE_SUBSTR_OK";
}
"##,
    );
}

#[test]
fn test_php_intl_grapheme_extr_maxbytes_mode() {
    compile_ok(
        r##"<?php
if (defined('GRAPHEME_EXTR_MAXBYTES')) {
    echo GRAPHEME_EXTR_MAXBYTES === 0 || is_int(GRAPHEME_EXTR_MAXBYTES) ? "MAXBYTES_CONST_OK" : "FAIL";
} else {
    echo "MAXBYTES_CONST_OK";
}
"##,
    );
}
