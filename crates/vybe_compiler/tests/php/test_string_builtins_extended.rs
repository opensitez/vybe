use super::helpers::compile_ok;

// ── str_word_count ───────────────────────────────────────────────
#[test]
fn str_word_count_sentence() {
    compile_ok(
        r#"<?php
$n = str_word_count("The quick brown fox");
echo $n;
"#,
    );
}

// ── wordwrap ─────────────────────────────────────────────────────
#[test]
fn wordwrap_word_boundary() {
    compile_ok(
        r#"<?php
$text = "The quick brown fox jumped over the lazy dog";
echo wordwrap($text, 15, "\n", false);
"#,
    );
}

// ── chunk_split ──────────────────────────────────────────────────
#[test]
fn chunk_split_with_separator() {
    compile_ok(
        r#"<?php
$encoded = chunk_split("abcdefghij", 3, "-");
echo $encoded;
"#,
    );
}

// ── number_format ────────────────────────────────────────────────
#[test]
fn number_format_thousands_decimals() {
    compile_ok(
        r#"<?php
echo number_format(9876543.21, 2, '.', ',');
echo number_format(1000);
echo number_format(0.125, 3);
"#,
    );
}

// ── vsprintf ─────────────────────────────────────────────────────
#[test]
fn vsprintf_array_args() {
    compile_ok(
        r#"<?php
$result = vsprintf("Name: %s, Score: %d, Ratio: %.2f", ["Alice", 95, 0.987]);
echo $result;
"#,
    );
}

// ── sscanf ───────────────────────────────────────────────────────
#[test]
fn sscanf_date_parsing() {
    compile_ok(
        r#"<?php
$parts = sscanf("2024-07-15", "%d-%d-%d");
echo $parts[0];
echo $parts[1];
echo $parts[2];
"#,
    );
}

// ── strip_tags ───────────────────────────────────────────────────
#[test]
fn strip_tags_allow_list() {
    compile_ok(
        r#"<?php
$html = "<h1>Title</h1><p>Body <b>text</b></p><script>alert(1)</script>";
echo strip_tags($html, "<h1><p>");
"#,
    );
}

// ── htmlspecialchars_decode ──────────────────────────────────────
#[test]
fn htmlspecialchars_decode_round_trip() {
    compile_ok(
        r#"<?php
$original = '<a href="url">link & text</a>';
$encoded = htmlspecialchars($original);
$decoded = htmlspecialchars_decode($encoded);
echo $decoded;
"#,
    );
}

// ── html_entity_decode ───────────────────────────────────────────
#[test]
fn html_entity_decode_named_entities() {
    compile_ok(
        r#"<?php
echo html_entity_decode("&lt;b&gt;bold&lt;/b&gt; &amp; &copy; &trade;");
"#,
    );
}

// ── addslashes ───────────────────────────────────────────────────
#[test]
fn addslashes_special_chars() {
    compile_ok(
        r#"<?php
$s = "He said 'hello' and \"goodbye\" with a \backslash";
echo addslashes($s);
"#,
    );
}

// ── stripslashes ─────────────────────────────────────────────────
#[test]
fn stripslashes_remove_escapes() {
    compile_ok(
        r#"<?php
$escaped = "It\'s a \\\"test\\\"";
echo stripslashes($escaped);
"#,
    );
}

// ── str_rot13 ────────────────────────────────────────────────────
#[test]
fn str_rot13_alphabet_rotation() {
    compile_ok(
        r#"<?php
$msg = "Hello World 123";
$rotated = str_rot13($msg);
echo $rotated;
echo str_rot13($rotated);
"#,
    );
}

// ── crc32 ────────────────────────────────────────────────────────
#[test]
fn crc32_checksum_consistency() {
    compile_ok(
        r#"<?php
$a = crc32("consistent input");
$b = crc32("consistent input");
echo ($a === $b) ? "stable" : "unstable";
echo is_int($a) ? "integer" : "not integer";
"#,
    );
}

// ── hex2bin ──────────────────────────────────────────────────────
#[test]
fn hex2bin_decode_hex_string() {
    compile_ok(
        r#"<?php
$binary = hex2bin("48656c6c6f");
echo $binary;
"#,
    );
}

// ── bin2hex ──────────────────────────────────────────────────────
#[test]
fn bin2hex_encode_bytes() {
    compile_ok(
        r#"<?php
$hex = bin2hex("Hi");
echo strtolower($hex);
echo strlen($hex);
"#,
    );
}

// ── strtr array form ─────────────────────────────────────────────
#[test]
fn strtr_array_substitution() {
    compile_ok(
        r#"<?php
$map = ["apple" => "fruit", "dog" => "animal", "blue" => "color"];
$result = strtr("apple and dog and blue", $map);
echo $result;
"#,
    );
}

// ── similar_text ─────────────────────────────────────────────────
#[test]
fn similar_text_comparison() {
    compile_ok(
        r#"<?php
$common = similar_text("World", "Word");
echo $common;
similar_text("Hello", "Hello", $pct);
echo ($pct == 100.0) ? "full" : "partial";
"#,
    );
}

// ── levenshtein ──────────────────────────────────────────────────
#[test]
fn levenshtein_edit_distance() {
    compile_ok(
        r#"<?php
echo levenshtein("kitten", "sitting");
echo levenshtein("sunday", "saturday");
echo levenshtein("abc", "abc");
"#,
    );
}

// ── soundex ──────────────────────────────────────────────────────
#[test]
fn soundex_phonetic_code() {
    compile_ok(
        r#"<?php
$code = soundex("Smith");
echo is_string($code) ? "ok" : "fail";
echo strlen($code) === 4 ? "four" : "other";
echo soundex("Smythe") === soundex("Smith") ? "match" : "no";
"#,
    );
}

// ── metaphone ────────────────────────────────────────────────────
#[test]
fn metaphone_phonetic_encoding() {
    compile_ok(
        r#"<?php
$m = metaphone("Thompson");
echo is_string($m) ? "ok" : "fail";
echo strlen($m) > 0 ? "nonempty" : "empty";
echo metaphone("Thomson") === metaphone("Tomson") ? "match" : "no";
"#,
    );
}

// ── quoted_printable_encode ──────────────────────────────────────
#[test]
fn quoted_printable_encode_basic() {
    compile_ok(
        r#"<?php
$encoded = quoted_printable_encode("Subject: =?UTF-8?");
echo is_string($encoded) ? "ok" : "fail";
$decoded = quoted_printable_decode($encoded);
echo is_string($decoded) ? "ok" : "fail";
"#,
    );
}

// ── str_getcsv ───────────────────────────────────────────────────
#[test]
fn str_getcsv_parse_csv_line() {
    compile_ok(
        r#"<?php
$fields = str_getcsv("one,two,three");
echo count($fields);
echo $fields[0];
echo $fields[2];
"#,
    );
}

// ── str_ireplace ─────────────────────────────────────────────────
#[test]
fn str_ireplace_case_insensitive_replace() {
    compile_ok(
        r#"<?php
$result = str_ireplace("HELLO", "Hi", "Hello World HELLO hello");
echo $result;
"#,
    );
}

// ── substr_count ─────────────────────────────────────────────────
#[test]
fn substr_count_occurrences() {
    compile_ok(
        r#"<?php
echo substr_count("hello world hello world", "hello");
echo substr_count("banana", "ana");
echo substr_count("aababab", "ab");
"#,
    );
}

// ── substr_replace ───────────────────────────────────────────────
#[test]
fn substr_replace_by_position() {
    compile_ok(
        r#"<?php
echo substr_replace("Hello World", "PHP", 6, 5);
echo substr_replace("abcdefgh", "XYZ", 2, 3);
echo substr_replace("insert here", ">>", 6, 0);
"#,
    );
}
