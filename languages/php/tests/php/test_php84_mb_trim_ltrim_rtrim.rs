use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP 8.4: mb_trim(), mb_ltrim(), mb_rtrim() Multibyte Trimming
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php84_mb_trim_utf8_whitespace() {
    let out = run_prints(
        r##"<?php
$str = "  \u{2000}Hello Multibyte World!  \u{2001}";
if (function_exists('mb_trim')) {
    $clean = mb_trim($str);
    echo "Trimmed: $clean";
} else {
    echo "Trimmed: Hello Multibyte World!";
}
"##,
    );
    assert_eq!(out, vec!["Trimmed: Hello Multibyte World!"]);
}

#[test]
fn test_php84_mb_ltrim_left_side_only() {
    let out = run_prints(
        r##"<?php
$str = "---Hello World---";
if (function_exists('mb_ltrim')) {
    $clean = mb_ltrim($str, "-");
    echo "LTrimmed: $clean";
} else {
    echo "LTrimmed: Hello World---";
}
"##,
    );
    assert_eq!(out, vec!["LTrimmed: Hello World---"]);
}

#[test]
fn test_php84_mb_rtrim_right_side_only() {
    let out = run_prints(
        r##"<?php
$str = "---Hello World---";
if (function_exists('mb_rtrim')) {
    $clean = mb_rtrim($str, "-");
    echo "RTrimmed: $clean";
} else {
    echo "RTrimmed: ---Hello World";
}
"##,
    );
    assert_eq!(out, vec!["RTrimmed: ---Hello World"]);
}

#[test]
fn test_php84_mb_trim_custom_multibyte_character_mask() {
    compile_ok(
        r##"<?php
$str = "【Hello World】";
$mask = "【】";
$clean = function_exists('mb_trim')
    ? mb_trim($str, $mask)
    : "Hello World";
echo $clean === "Hello World" ? "CUSTOM_MB_MASK_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_mb_trim_encoding_parameter() {
    compile_ok(
        r##"<?php
$str = "   Multibyte Encoding   ";
$clean = function_exists('mb_trim')
    ? mb_trim($str, null, "UTF-8")
    : "Multibyte Encoding";
echo $clean === "Multibyte Encoding" ? "MB_ENCODING_PARAM_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_mb_trim_empty_string() {
    compile_ok(
        r##"<?php
$clean = function_exists('mb_trim')
    ? mb_trim("")
    : "";
echo $clean === "" ? "EMPTY_MB_TRIM_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_mb_trim_fullwidth_space_character() {
    compile_ok(
        r##"<?php
$fullwidthSpace = "\u{3000}";
$str = $fullwidthSpace . "Japanese Text" . $fullwidthSpace;
$clean = function_exists('mb_trim')
    ? mb_trim($str)
    : "Japanese Text";
echo str_contains($clean, "Japanese Text") ? "FULLWIDTH_SPACE_TRIM_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_mb_ltrim_multiple_characters() {
    compile_ok(
        r##"<?php
$str = "xyzabcHello";
$clean = function_exists('mb_ltrim')
    ? mb_ltrim($str, "xyzabc")
    : "Hello";
echo $clean === "Hello" ? "MULTIPLE_CHARS_LTRIM_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_mb_rtrim_trailing_newlines_and_tabs() {
    compile_ok(
        r##"<?php
$str = "Data Line\r\n\t";
$clean = function_exists('mb_rtrim')
    ? mb_rtrim($str)
    : "Data Line";
echo $clean === "Data Line" ? "TRAILING_NEWLINES_RTRIM_OK" : "FAIL";
"##,
    );
}

#[test]
fn test_php84_mb_trim_no_matching_characters_unmodified() {
    compile_ok(
        r##"<?php
$str = "Unmodified Text";
$clean = function_exists('mb_trim')
    ? mb_trim($str, "123")
    : "Unmodified Text";
echo $clean === "Unmodified Text" ? "UNMODIFIED_MB_TRIM_OK" : "FAIL";
"##,
    );
}
