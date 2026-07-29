use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: String Manipulation & Formatting — substr, str_replace, explode, implode, sprintf, str_contains, str_starts_with, str_ends_with
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_string_substr_positive_negative() {
    let out = run_prints(
        r#"<?php
$str = "Hello World";
echo substr($str, 0, 5) . " | " . substr($str, -5);
"#,
    );
    assert_eq!(out, vec!["Hello | World"]);
}

#[test]
fn test_php_string_str_replace_array_replacement() {
    let out = run_prints(
        r#"<?php
$vowels = ["a", "e", "i", "o", "u"];
$res = str_replace($vowels, "*", "Hello World");
echo $res;
"#,
    );
    assert_eq!(out, vec!["H*ll* W*rld"]);
}

#[test]
fn test_php_string_explode_implode_delimited() {
    let out = run_prints(
        r#"<?php
$csv = "apple,banana,cherry";
$parts = explode(",", $csv);
echo implode(" - ", $parts);
"#,
    );
    assert_eq!(out, vec!["apple - banana - cherry"]);
}

#[test]
fn test_php_string_trim_ltrim_rtrim_custom_characters() {
    let out = run_prints(
        r#"<?php
$raw = "///Hello World///";
echo trim($raw, "/") . " | " . ltrim($raw, "/");
"#,
    );
    assert_eq!(out, vec!["Hello World | Hello World///"]);
}

#[test]
fn test_php_string_sprintf_format_specifiers() {
    let out = run_prints(
        r#"<?php
$formatted = sprintf("User %s ID %04d Price $%.2f", "Alice", 42, 19.95);
echo $formatted;
"#,
    );
    assert_eq!(out, vec!["User Alice ID 0042 Price $19.95"]);
}

#[test]
fn test_php_string_php8_str_contains_starts_ends() {
    let out = run_prints(
        r#"<?php
$haystack = "https://laravel.com/docs";
echo str_starts_with($haystack, "https://") ? "YES" : "NO";
echo " ";
echo str_ends_with($haystack, "/docs") ? "YES" : "NO";
echo " ";
echo str_contains($haystack, "laravel") ? "YES" : "NO";
"#,
    );
    assert_eq!(out, vec!["YES YES YES"]);
}

#[test]
fn test_php_string_pad_left_right_both() {
    let out = run_prints(
        r#"<?php
$num = "123";
echo str_pad($num, 6, "0", STR_PAD_LEFT);
"#,
    );
    assert_eq!(out, vec!["000123"]);
}

#[test]
fn test_php_string_repeat_multiplier() {
    let out = run_prints(
        r#"<?php
echo str_repeat("Na", 3) . " Batman";
"#,
    );
    assert_eq!(out, vec!["NaNaNa Batman"]);
}

#[test]
fn test_php_string_strpos_stripos_offset() {
    let out = run_prints(
        r#"<?php
$text = "The quick brown fox jumps over the lazy dog";
$pos = strpos($text, "brown");
echo $pos !== false ? "FOUND_$pos" : "NOT_FOUND";
"#,
    );
    assert_eq!(out, vec!["FOUND_10"]);
}

#[test]
fn test_php_string_wordwrap_breaking() {
    compile_ok(
        r#"<?php
$text = "A very long words sentence that needs wrapping";
$newtext = wordwrap($text, 10, "\n", true);
echo $newtext;
"#,
    );
}

#[test]
fn test_php_string_strtok_tokenizer() {
    compile_ok(
        r#"<?php
$string = "This is\tan example\nstring";
$tok = strtok($string, " \n\t");
while ($tok !== false) {
    echo "Word=$tok\n";
    $tok = strtok(" \n\t");
}
"#,
    );
}

#[test]
fn test_php_string_number_format_decimals() {
    compile_ok(
        r#"<?php
$number = 1234.5678;
echo number_format($number, 2, '.', ',');
"#,
    );
}

#[test]
fn test_php_string_parse_str_query() {
    compile_ok(
        r#"<?php
$str = "first=value&arr[]=foo+bar&arr[]=baz";
parse_str($str, $output);
echo $output['first'];
"#,
    );
}

#[test]
fn test_php_string_str_tr_character_translation() {
    compile_ok(
        r#"<?php
$trans = ["h" => "hello", "hello" => "hi"];
echo strtr("hi all, I said hello", $trans);
"#,
    );
}

#[test]
fn test_php_string_addslashes_stripslashes() {
    compile_ok(
        r#"<?php
$str = "Is your name O'Reilly?";
$escaped = addslashes($str);
echo stripslashes($escaped);
"#,
    );
}

#[test]
fn test_php_string_implode_preserves_keyed_order_for_values() {
    let out = run_prints(
        r#"<?php
$items = [10 => 'ten', '2' => 'two', 1 => 'one', 3 => 'three'];
echo implode('|', array_values($items)) . '|' . count($items);
"#,
    );
    assert_eq!(out, vec!["ten|two|one|three|4"]);
}

#[test]
fn test_php_string_explode_limit_and_empty_tail() {
    let out = run_prints(
        r#"<?php
$parts = explode(':', 'a:b:c:', 3);
echo count($parts);
echo '|';
echo $parts[0];
echo '|';
echo $parts[1];
echo '|';
echo $parts[2];
"#,
    );
    assert_eq!(out, vec!["3|a|b|c:"]);
}

#[test]
fn test_php_string_implode_empty_glue_with_arrays() {
    let out = run_prints(
        r#"<?php
$parts = ['a', 'b', 'c'];
echo implode('', $parts) . '|' . strlen(implode('', $parts));
"#,
    );
    assert_eq!(out, vec!["abc|3"]);
}

#[test]
fn test_php_string_explode_array_and_join_unicode() {
    let out = run_prints(
        r#"<?php
$text = "a,b,,c,,";
$items = explode(',', $text);
$joined = implode('|', $items);
echo count($items) . '|' . substr($joined, 0, 12);
"#,
    );
    assert_eq!(out, vec!["6|a|b||c||"]);
}

#[test]
fn test_php_string_strpos_offset_edge_cases() {
    let out = run_prints(
        r#"<?php
echo strpos("hello", "h", 0);
echo "|";
echo strpos("hello", "h", -1) === 0 ? "ok" : "bad";
echo "|";
echo strpos("hello", "l", 2);
"#,
    );
    assert_eq!(out, vec!["0|ok|3"]);
}

#[test]
fn test_php_string_strrchr_and_strstr_mix() {
    let out = run_prints(
        r#"<?php
echo strrchr("ababa", "ba");
echo "|";
echo strstr("ababa", "ba");
"#,
    );
    assert_eq!(out, vec!["ba|baba"]);
}

#[test]
fn test_php_string_chunk_split_length_zero_and_padding() {
    let out = run_prints(
        r#"<?php
echo chunk_split("abcdef", 3, "->");
echo "|";
echo str_replace("e", "", "seed", $count);
echo "|";
echo $count;
"#,
    );
    assert_eq!(out, vec!["abc->def->|sd|1"]);
}

#[test]
fn test_php_string_sprintf_zero_precision_and_width() {
    let out = run_prints(
        r#"<?php
echo sprintf("%'.8s", "abc");
echo "|";
echo sprintf("%.0f", 2.49);
echo "|";
echo sprintf("%+ .3f", 7.5);
"#,
    );
    assert_eq!(out, vec![".....abc|2|+7.500"]);
}
