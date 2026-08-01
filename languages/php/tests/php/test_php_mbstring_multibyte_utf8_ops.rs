use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Multibyte String Operations — mb_strlen, mb_substr, mb_strpos, mb_strtolower, mb_strtoupper, mb_convert_encoding, mb_check_encoding
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_mb_strlen_vs_strlen_multibyte() {
    let out = run_prints(
        r#"<?php
$str = "Héllo Wörld €";
echo strlen($str) . " vs " . mb_strlen($str, "UTF-8");
"#,
    );
    assert_eq!(out, vec!["17 vs 13"]);
}

#[test]
fn test_php_mb_substr_unicode_slice() {
    let out = run_prints(
        r#"<?php
$str = "こんにちは世界"; // Hello World in Japanese
echo mb_substr($str, 0, 5, "UTF-8");
"#,
    );
    assert_eq!(out, vec!["こんにちは"]);
}

#[test]
fn test_php_mb_strtoupper_mb_strtolower_case_folding() {
    let out = run_prints(
        r#"<?php
$str = "münchen";
echo mb_strtoupper($str, "UTF-8") . " | " . mb_strtolower("MÜNCHEN", "UTF-8");
"#,
    );
    assert_eq!(out, vec!["MÜNCHEN | münchen"]);
}

#[test]
fn test_php_mb_strpos_character_offset() {
    let out = run_prints(
        r#"<?php
$str = "αβγδεζηθ";
echo mb_strpos($str, "δε", 0, "UTF-8");
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn test_php_mb_check_encoding_validation() {
    let out = run_prints(
        r#"<?php
$validUtf8 = "Valid UTF-8 string 🔥";
echo mb_check_encoding($validUtf8, "UTF-8") ? "VALID" : "INVALID";
"#,
    );
    assert_eq!(out, vec!["VALID"]);
}

#[test]
fn test_php_mb_str_split_chunking() {
    compile_ok(
        r#"<?php
$str = "你好世界";
$chars = mb_str_split($str, 1, "UTF-8");
echo implode("-", $chars);
"#,
    );
}

#[test]
fn test_php_mb_convert_encoding_conversion() {
    compile_ok(
        r#"<?php
$utf8 = "Test string";
$iso = mb_convert_encoding($utf8, "ISO-8859-1", "UTF-8");
echo strlen($iso) > 0 ? "CONVERTED" : "FAIL";
"#,
    );
}

#[test]
fn test_php_mb_detect_encoding_auto() {
    compile_ok(
        r#"<?php
$text = "Simple ASCII string";
$encoding = mb_detect_encoding($text, ["UTF-8", "ASCII", "ISO-8859-1"]);
echo $encoding;
"#,
    );
}

#[test]
fn test_php_mb_substr_count_occurrences() {
    compile_ok(
        r#"<?php
$haystack = "a-b-c-a-b-a";
echo mb_substr_count($haystack, "a", "UTF-8");
"#,
    );
}

#[test]
fn test_php_mb_scrub_invalid_encoding() {
    compile_ok(
        r#"<?php
$invalid = "Hello \xFF\xFE World";
$scrubbed = mb_scrub($invalid, "UTF-8");
echo strlen($scrubbed) > 0 ? "SCRUBBED" : "EMPTY";
"#,
    );
}

#[test]
fn test_php_mb_strlen_empty_and_zero_offset_runtime() {
    let out = run_prints(
        r#"<?php
echo mb_strlen("", "UTF-8");
echo "|";
echo mb_strlen("abc", "UTF-8");
"#,
    );
    assert_eq!(out, vec!["0|3"]);
}

#[test]
fn test_php_mb_substr_start_beyond_length_runtime() {
    let out = run_prints(
        r#"<?php
echo mb_substr("世界", 10, 2, "UTF-8") === "" ? "empty" : "non-empty";
echo "|";
echo mb_substr("世界", -10, 1, "UTF-8");
"#,
    );
    assert_eq!(out, vec!["empty|界"]);
}

#[test]
fn test_php_mb_strrpos_case_fold_runtime() {
    let out = run_prints(
        r#"<?php
echo mb_strripos("aĀbĀc", "ā");
echo "|";
echo mb_strrpos("a-b-c-a", "a", 0);
"#,
    );
    assert_eq!(out, vec!["3|6"]);
}

#[test]
fn test_php_mb_strstr_needle_not_found_runtime() {
    let out = run_prints(
        r#"<?php
echo mb_strstr("abcdef", "z") === false ? "missing" : "found";
echo "|";
echo mb_strrchr("caféc", "é") === "éc" ? "tail" : "no";
"#,
    );
    assert_eq!(out, vec!["missing|tail"]);
}

// ── Astral (non-BMP) code points ───────────────────────────────────────
//
// `mb_*` counts CODE POINTS; `strlen`/`substr` count BYTES. Outside the BMP
// one code point is TWO UTF-16 units, which is where both units diverge from
// the UTF-16 count the emitters used to produce — `mb_substr` returned half a
// surrogate pair. Every value below is measured against real `php`.
//
// Each case is written TWICE, once with a literal receiver and once through a
// function so the value is not constant-foldable. The walker folds `mb_strlen`
// on any expression `lit_str_arg` can resolve — including a simple variable —
// so a literal-only test passes without the emitter being reached at all.

#[test]
fn test_php_mb_astral_literal_receiver() {
    let out = run_prints(
        r#"<?php
echo mb_strlen("a😀b€c") . " " . mb_substr("a😀b€c", 1, 1) . " " . mb_strpos("a😀b€c", "€") . " " . implode(",", mb_str_split("a😀b"));
"#,
    );
    assert_eq!(out, vec!["5 😀 3 a,😀,b"]);
}

#[test]
fn test_php_mb_astral_dynamic_receiver() {
    let out = run_prints(
        r#"<?php
function f($x) { return $x; }
$s = f("a😀b€c");
echo mb_strlen($s) . " " . mb_substr($s, 1, 1) . " " . mb_substr($s, 1, 2) . " " . mb_strpos($s, "€") . " " . mb_substr($s, -2) . " " . strlen($s);
"#,
    );
    assert_eq!(out, vec!["5 😀 😀b 3 €c 10"]);
}

/// `mb_strpos` answers `false`, not `-1`, when the needle is absent — and the
/// offset argument counts in the same unit as the result.
#[test]
fn test_php_mb_strpos_absent_and_offset() {
    let out = run_prints(
        r#"<?php
function f($x) { return $x; }
$s = f("a😀b😀c");
var_dump(mb_strpos($s, "z"));
echo mb_strpos($s, "😀", 2);
"#,
    );
    assert_eq!(out, vec!["bool(false)", "3"]);
}
