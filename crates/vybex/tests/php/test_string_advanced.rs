use super::helpers::run_prints;

// ── substr_count / substr_replace ────────────────────────────────
#[test]
fn substr_count_basic() {
    assert_eq!(run_prints(r#"<?php
echo substr_count("hello world hello", "hello");
echo substr_count("banana", "an");
"#), &["2", "2"]);
}

#[test]
fn substr_replace_basic() {
    assert_eq!(run_prints(r#"<?php
echo substr_replace("hello world", "PHP", 6, 5);
echo substr_replace("abcdef", "XY", 2, 2);
"#), &["hello PHP", "abXYef"]);
}

// ── str_split ────────────────────────────────────────────────────
#[test]
fn str_split_basic() {
    assert_eq!(run_prints(r#"<?php
$chars = str_split("hello");
echo implode(",", $chars);
$chunks = str_split("abcdefgh", 3);
echo implode("|", $chunks);
"#), &["h,e,l,l,o", "abc|def|gh"]);
}

// ── str_word_count ───────────────────────────────────────────────
#[test]
fn str_word_count_basic() {
    assert_eq!(run_prints(r#"<?php
echo str_word_count("Hello beautiful world");
"#), &["3"]);
}

// ── wordwrap ─────────────────────────────────────────────────────
#[test]
fn wordwrap_basic() {
    assert_eq!(run_prints(r#"<?php
$text = "The quick brown fox jumped over the lazy dog";
$wrapped = wordwrap($text, 15, "\n", true);
$lines = explode("\n", $wrapped);
echo count($lines);
"#), &["4"]);
}

// ── number_format ────────────────────────────────────────────────
#[test]
fn number_format_basic() {
    assert_eq!(run_prints(r#"<?php
echo number_format(1234567.891);
echo number_format(1234567.891, 2);
echo number_format(1234567.891, 2, ",", ".");
"#), &["1,234,568", "1,234,567.89", "1.234.567,89"]);
}

#[test]
fn number_format_small() {
    assert_eq!(run_prints(r#"<?php
echo number_format(0.5, 0);
echo number_format(42, 3);
echo number_format(1000, 0, ".", ",");
"#), &["1", "42.000", "1,000"]);
}

// ── similar_text / levenshtein ───────────────────────────────────
#[test]
fn similar_text_basic() {
    assert_eq!(run_prints(r#"<?php
echo similar_text("Hello", "World");
echo similar_text("abc", "abc");
"#), &["2", "3"]);
}

#[test]
fn levenshtein_basic() {
    assert_eq!(run_prints(r#"<?php
echo levenshtein("kitten", "sitting");
echo levenshtein("hello", "hello");
echo levenshtein("", "abc");
"#), &["3", "0", "3"]);
}

// ── soundex / metaphone ──────────────────────────────────────────
#[test]
fn soundex_basic() {
    assert_eq!(run_prints(r#"<?php
echo soundex("Robert");
echo soundex("Rupert");
echo soundex("Robert") == soundex("Rupert") ? "match" : "no match";
"#), &["R163", "R163", "match"]);
}

#[test]
fn metaphone_basic() {
    assert_eq!(run_prints(r#"<?php
echo metaphone("Thompson");
echo metaphone("Thomson");
"#), &["TMPSN", "TMSN"]);
}

// ── chunk_split ──────────────────────────────────────────────────
#[test]
fn chunk_split_basic() {
    assert_eq!(run_prints(r#"<?php
$result = chunk_split("abcdefgh", 3, "-");
echo $result;
"#), &["abc-def-gh-"]);
}

// ── str_getcsv ───────────────────────────────────────────────────
#[test]
fn str_getcsv_basic() {
    assert_eq!(run_prints(r#"<?php
$fields = str_getcsv("one,two,three");
echo implode("|", $fields);
$quoted = str_getcsv('"hello, world","test"');
echo implode("|", $quoted);
"#), &["one|two|three", "hello, world|test"]);
}

// ── String padding / repeat (extended) ───────────────────────────
#[test]
fn str_pad_all_modes() {
    assert_eq!(run_prints(r#"<?php
echo str_pad("42", 5, "0", STR_PAD_LEFT);
echo str_pad("hi", 10, "-");
echo str_pad("x", 5, "AB", STR_PAD_BOTH);
"#), &["00042", "hi--------", "ABxAB"]);
}

// ── mb_ multibyte functions ──────────────────────────────────────
#[test]
fn mb_strlen_basic() {
    assert_eq!(run_prints(r#"<?php
echo mb_strlen("hello");
echo mb_strlen("こんにちは");
"#), &["5", "5"]);
}

#[test]
fn mb_strtoupper_lower() {
    assert_eq!(run_prints(r#"<?php
echo mb_strtoupper("hello");
echo mb_strtolower("WORLD");
"#), &["HELLO", "world"]);
}

#[test]
fn mb_substr_basic() {
    assert_eq!(run_prints(r#"<?php
echo mb_substr("hello world", 6);
echo mb_substr("hello", 0, 3);
"#), &["world", "hel"]);
}

// ── ctype_ functions ─────────────────────────────────────────────
#[test]
fn ctype_alpha_digit() {
    assert_eq!(run_prints(r#"<?php
echo ctype_alpha("hello") ? "yes" : "no";
echo ctype_alpha("hello123") ? "yes" : "no";
echo ctype_digit("12345") ? "yes" : "no";
echo ctype_digit("123a5") ? "yes" : "no";
"#), &["yes", "no", "yes", "no"]);
}

#[test]
fn ctype_alnum_space() {
    assert_eq!(run_prints(r#"<?php
echo ctype_alnum("hello123") ? "yes" : "no";
echo ctype_space("   \t\n") ? "yes" : "no";
echo ctype_upper("ABC") ? "yes" : "no";
echo ctype_lower("abc") ? "yes" : "no";
"#), &["yes", "yes", "yes", "yes"]);
}

// ── String conversion functions ──────────────────────────────────
#[test]
fn bin2hex_hex2bin() {
    assert_eq!(run_prints(r#"<?php
$hex = bin2hex("AB");
echo $hex;
echo hex2bin($hex);
"#), &["4142", "AB"]);
}

#[test]
fn base64_encode_decode() {
    assert_eq!(run_prints(r#"<?php
$encoded = base64_encode("Hello PHP");
echo $encoded;
echo base64_decode($encoded);
"#), &["SGVsbG8gUEhQ", "Hello PHP"]);
}

#[test]
fn urlencode_decode() {
    assert_eq!(run_prints(r#"<?php
$encoded = urlencode("hello world&foo=bar");
echo $encoded;
echo urldecode($encoded);
"#), &["hello+world%26foo%3Dbar", "hello world&foo=bar"]);
}

// ── String search functions ──────────────────────────────────────
#[test]
fn strstr_basic() {
    assert_eq!(run_prints(r#"<?php
echo strstr("user@example.com", "@");
echo strstr("user@example.com", "@", true);
"#), &["@example.com", "user"]);
}

#[test]
fn strrpos_basic() {
    assert_eq!(run_prints(r#"<?php
echo strrpos("hello world hello", "hello");
echo strrpos("abcabc", "bc");
"#), &["12", "4"]);
}

// ── String manipulation ──────────────────────────────────────────
#[test]
fn str_replace_array() {
    assert_eq!(run_prints(r#"<?php
$result = str_replace(["a", "e", "i", "o", "u"], "*", "hello world");
echo $result;
"#), &["h*ll* w*rld"]);
}

#[test]
fn str_ireplace_case_insensitive() {
    assert_eq!(run_prints(r#"<?php
echo str_ireplace("HELLO", "hi", "Hello World hello");
"#), &["hi World hi"]);
}

#[test]
fn string_reverse() {
    assert_eq!(run_prints(r#"<?php
echo strrev("hello");
echo strrev("12345");
"#), &["olleh", "54321"]);
}

#[test]
fn string_ucwords() {
    assert_eq!(run_prints(r#"<?php
echo ucwords("hello beautiful world");
echo ucwords("one-two-three", "-");
"#), &["Hello Beautiful World", "One-Two-Three"]);
}

#[test]
fn sprintf_advanced() {
    assert_eq!(run_prints(r#"<?php
echo sprintf("%05d", 42);
echo sprintf("%.2f", 3.14159);
echo sprintf("%s has %d items", "cart", 5);
echo sprintf("%10s", "right");
echo sprintf("%-10s|", "left");
"#), &["00042", "3.14", "cart has 5 items", "     right", "left      |"]);
}
