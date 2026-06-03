use super::helpers::run_prints;

// ── substr_count / substr_replace ────────────────────────────────
#[test]
fn substr_count_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo substr_count("hello world hello", "hello");
echo substr_count("banana", "an");
"#
        ),
        &["2", "2"]
    );
}

#[test]
fn substr_replace_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo substr_replace("hello world", "PHP", 6, 5);
echo substr_replace("abcdef", "XY", 2, 2);
"#
        ),
        &["hello PHP", "abXYef"]
    );
}

// ── str_split ────────────────────────────────────────────────────
#[test]
fn str_split_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$chars = str_split("hello");
echo implode(",", $chars);
$chunks = str_split("abcdefgh", 3);
echo implode("|", $chunks);
"#
        ),
        &["h,e,l,l,o", "abc|def|gh"]
    );
}

// ── wordwrap ─────────────────────────────────────────────────────
#[test]
fn wordwrap_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$text = "The quick brown fox jumped over the lazy dog";
$wrapped = wordwrap($text, 15, "\n", true);
$lines = explode("\n", $wrapped);
echo count($lines);
"#
        ),
        &["4"]
    );
}

// ── number_format ────────────────────────────────────────────────
#[test]
fn number_format_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo number_format(1234567.891);
echo number_format(1234567.891, 2);
echo number_format(1234567.891, 2, ",", ".");
"#
        ),
        &["1,234,568", "1,234,567.89", "1.234.567,89"]
    );
}

#[test]
fn number_format_small() {
    assert_eq!(
        run_prints(
            r#"<?php
echo number_format(0.5, 0);
echo number_format(42, 3);
echo number_format(1000, 0, ".", ",");
"#
        ),
        &["1", "42.000", "1,000"]
    );
}

// ── similar_text / levenshtein ───────────────────────────────────
#[test]
fn similar_text_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo similar_text("Hello", "World");
echo similar_text("abc", "abc");
"#
        ),
        &["2", "3"]
    );
}

#[test]
fn levenshtein_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo levenshtein("kitten", "sitting");
echo levenshtein("hello", "hello");
echo levenshtein("", "abc");
"#
        ),
        &["3", "0", "3"]
    );
}

// ── soundex / metaphone ──────────────────────────────────────────
#[test]
fn soundex_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo soundex("Robert");
echo soundex("Rupert");
echo soundex("Robert") == soundex("Rupert") ? "match" : "no match";
"#
        ),
        &["R163", "R163", "match"]
    );
}

#[test]
fn metaphone_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo metaphone("Thompson");
echo metaphone("Thomson");
"#
        ),
        &["TMPSN", "TMSN"]
    );
}

// ── chunk_split ──────────────────────────────────────────────────
#[test]
fn chunk_split_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$result = chunk_split("abcdefgh", 3, "-");
echo $result;
"#
        ),
        &["abc-def-gh-"]
    );
}

// ── str_getcsv ───────────────────────────────────────────────────
#[test]
fn str_getcsv_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$fields = str_getcsv("one,two,three");
echo implode("|", $fields);
$quoted = str_getcsv('"hello, world","test"');
echo implode("|", $quoted);
"#
        ),
        &["one|two|three", "hello, world|test"]
    );
}

// ── String padding / repeat (extended) ───────────────────────────
#[test]
fn str_pad_all_modes() {
    assert_eq!(
        run_prints(
            r#"<?php
echo str_pad("42", 5, "0", STR_PAD_LEFT);
echo str_pad("hi", 10, "-");
echo str_pad("x", 5, "AB", STR_PAD_BOTH);
"#
        ),
        &["00042", "hi--------", "ABxAB"]
    );
}

// ── ctype_ functions ─────────────────────────────────────────────
#[test]
fn ctype_alpha_digit() {
    assert_eq!(
        run_prints(
            r#"<?php
echo ctype_alpha("hello") ? "yes" : "no";
echo ctype_alpha("hello123") ? "yes" : "no";
echo ctype_digit("12345") ? "yes" : "no";
echo ctype_digit("123a5") ? "yes" : "no";
"#
        ),
        &["yes", "no", "yes", "no"]
    );
}

// ── String conversion functions ──────────────────────────────────
#[test]
fn bin2hex_hex2bin() {
    assert_eq!(
        run_prints(
            r#"<?php
$hex = bin2hex("AB");
echo $hex;
echo hex2bin($hex);
"#
        ),
        &["4142", "AB"]
    );
}

#[test]
fn base64_encode_decode() {
    assert_eq!(
        run_prints(
            r#"<?php
$encoded = base64_encode("Hello PHP");
echo $encoded;
echo base64_decode($encoded);
"#
        ),
        &["SGVsbG8gUEhQ", "Hello PHP"]
    );
}

#[test]
fn urlencode_decode() {
    assert_eq!(
        run_prints(
            r#"<?php
$encoded = urlencode("hello world&foo=bar");
echo $encoded;
echo urldecode($encoded);
"#
        ),
        &["hello+world%26foo%3Dbar", "hello world&foo=bar"]
    );
}

// ── String search functions ──────────────────────────────────────
#[test]
fn strstr_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strstr("user@example.com", "@");
echo strstr("user@example.com", "@", true);
"#
        ),
        &["@example.com", "user"]
    );
}

#[test]
fn strrpos_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strrpos("hello world hello", "hello");
echo strrpos("abcabc", "bc");
"#
        ),
        &["12", "4"]
    );
}

// ── String manipulation ──────────────────────────────────────────
#[test]
fn str_replace_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$result = str_replace(["a", "e", "i", "o", "u"], "*", "hello world");
echo $result;
"#
        ),
        &["h*ll* w*rld"]
    );
}

#[test]
fn str_ireplace_case_insensitive() {
    assert_eq!(
        run_prints(
            r#"<?php
echo str_ireplace("HELLO", "hi", "Hello World hello");
"#
        ),
        &["hi World hi"]
    );
}

#[test]
fn string_reverse() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strrev("hello");
echo strrev("12345");
"#
        ),
        &["olleh", "54321"]
    );
}

#[test]
fn string_ucwords() {
    assert_eq!(
        run_prints(
            r#"<?php
echo ucwords("hello beautiful world");
echo ucwords("one-two-three", "-");
"#
        ),
        &["Hello Beautiful World", "One-Two-Three"]
    );
}

#[test]
fn sprintf_advanced() {
    assert_eq!(
        run_prints(
            r#"<?php
echo sprintf("%05d", 42);
echo sprintf("%.2f", 3.14159);
echo sprintf("%s has %d items", "cart", 5);
echo sprintf("%10s", "right");
echo sprintf("%-10s|", "left");
"#
        ),
        &[
            "00042",
            "3.14",
            "cart has 5 items",
            "     right",
            "left      |"
        ]
    );
}

// ── sprintf hex / octal / binary / scientific ────────────────────
#[test]
fn sprintf_format_modes() {
    assert_eq!(
        run_prints(
            r#"<?php
echo sprintf("%x", 255);
echo sprintf("%X", 255);
echo sprintf("%o", 8);
echo sprintf("%b", 10);
echo sprintf("%e", 123456.789);
"#
        ),
        &["ff", "FF", "10", "1010", "1.234568e+5"]
    );
}

// ── vsprintf ─────────────────────────────────────────────────────
#[test]
fn vsprintf_with_array() {
    assert_eq!(
        run_prints(
            r#"<?php
$args = ["Alice", 30, "NYC"];
echo vsprintf("%s is %d years old and lives in %s", $args);
"#
        ),
        &["Alice is 30 years old and lives in NYC"]
    );
}

// ── str_contains / str_starts_with / str_ends_with ───────────────
#[test]
fn str_contains_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo str_contains("Hello World", "World") ? "yes" : "no";
echo str_contains("Hello World", "world") ? "yes" : "no";
echo str_contains("", "") ? "yes" : "no";
"#
        ),
        &["yes", "no", "yes"]
    );
}

#[test]
fn str_starts_ends_with_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo str_starts_with("Hello World", "Hello") ? "yes" : "no";
echo str_starts_with("Hello World", "World") ? "yes" : "no";
echo str_ends_with("Hello World", "World") ? "yes" : "no";
echo str_ends_with("Hello World", "Hello") ? "yes" : "no";
"#
        ),
        &["yes", "no", "yes", "no"]
    );
}

// ── strtr with array map ─────────────────────────────────────────
#[test]
fn strtr_array_map() {
    assert_eq!(
        run_prints(
            r#"<?php
$map = ["Hello" => "Hi", "World" => "Earth"];
echo strtr("Hello World", $map);
"#
        ),
        &["Hi Earth"]
    );
}

// ── str_rot13 ────────────────────────────────────────────────────
#[test]
fn str_rot13_roundtrip() {
    assert_eq!(
        run_prints(
            r#"<?php
$original = "Hello World";
$rotated = str_rot13($original);
echo $rotated;
echo str_rot13($rotated);
"#
        ),
        &["Uryyb Jbeyq", "Hello World"]
    );
}

// ── stripos case-insensitive position ────────────────────────────
#[test]
fn stripos_case_insensitive() {
    assert_eq!(
        run_prints(
            r#"<?php
echo stripos("Hello World", "WORLD");
echo stripos("PHP is great", "IS");
echo (stripos("no match", "XYZ") === false) ? "not found" : "found";
"#
        ),
        &["6", "4", "not found"]
    );
}

// ── strrchr last occurrence ──────────────────────────────────────
#[test]
fn strrchr_last_occurrence() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strrchr("/var/www/html/index.php", "/");
echo strrchr("user@example.com", "@");
"#
        ),
        &["/index.php", "@example.com"]
    );
}

// ── substr with negative offset ──────────────────────────────────
#[test]
fn substr_negative_offset() {
    assert_eq!(
        run_prints(
            r#"<?php
echo substr("Hello World", -5);
echo substr("Hello World", -5, 3);
echo substr("abcdef", 0, -2);
"#
        ),
        &["World", "Wor", "abcd"]
    );
}

// ── trim with custom chars ────────────────────────────────────────
#[test]
fn trim_custom_chars() {
    assert_eq!(
        run_prints(
            r#"<?php
echo trim("***hello***", "*");
echo trim("/path/to/file/", "/");
echo ltrim("000123", "0");
echo rtrim("hello...", ".");
"#
        ),
        &["hello", "path/to/file", "123", "hello"]
    );
}

// ── htmlspecialchars_decode ──────────────────────────────────────
#[test]
fn htmlspecialchars_decode_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$encoded = "&lt;div class=&quot;test&quot;&gt;Hello &amp; World&lt;/div&gt;";
echo htmlspecialchars_decode($encoded);
"#
        ),
        &["<div class=\"test\">Hello & World</div>"]
    );
}

// ── strip_tags ───────────────────────────────────────────────────
#[test]
fn strip_tags_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strip_tags("<p>Hello <b>World</b></p>");
echo strip_tags("<a href='url'>click</a> here", "<a>");
"#
        ),
        &["Hello World", "<a href='url'>click</a> here"]
    );
}

// ── addslashes / stripslashes ────────────────────────────────────
#[test]
fn addslashes_stripslashes() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = "It's a \"test\" with \\backslash";
$slashed = addslashes($s);
echo $slashed;
echo stripslashes($slashed);
"#
        ),
        &[
            "It\\'s a \\\"test\\\" with \\\\backslash",
            "It's a \"test\" with \\backslash"
        ]
    );
}

// ── crc32 ────────────────────────────────────────────────────────
#[test]
fn crc32_deterministic() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = crc32("hello");
$b = crc32("hello");
echo ($a === $b) ? "same" : "diff";
echo ($a !== crc32("world")) ? "unique" : "collision";
"#
        ),
        &["same", "unique"]
    );
}

// ── md5 / sha1 length checks ─────────────────────────────────────
#[test]
fn md5_sha1_lengths() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strlen(md5("hello"));
echo strlen(sha1("hello"));
echo md5("hello") === md5("hello") ? "stable" : "unstable";
"#
        ),
        &["32", "40", "stable"]
    );
}

// ── strcmp / strcasecmp ──────────────────────────────────────────
#[test]
fn strcmp_strcasecmp() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strcmp("abc", "abc") === 0 ? "equal" : "not equal";
echo strcmp("abc", "abd") < 0 ? "less" : "not less";
echo strcasecmp("Hello", "hello") === 0 ? "equal" : "not equal";
echo strcasecmp("ABC", "xyz") < 0 ? "less" : "not less";
"#
        ),
        &["equal", "less", "equal", "less"]
    );
}

// ── implode on single-element array ─────────────────────────────
#[test]
fn implode_single_element() {
    assert_eq!(
        run_prints(
            r#"<?php
echo implode(",", ["only"]);
echo implode("|", [42]);
echo implode("", ["x"]);
"#
        ),
        &["only", "42", "x"]
    );
}

// ── explode with limit ───────────────────────────────────────────
#[test]
fn explode_with_limit() {
    assert_eq!(
        run_prints(
            r#"<?php
$parts = explode(",", "a,b,c,d,e", 3);
echo count($parts);
echo $parts[0];
echo $parts[2];
"#
        ),
        &["3", "a", "c,d,e"]
    );
}

// ── str_repeat runtime ───────────────────────────────────────────
#[test]
fn str_repeat_runtime_assertion() {
    assert_eq!(
        run_prints(
            r#"<?php
echo str_repeat("ab", 4);
echo str_repeat("-", 5);
echo strlen(str_repeat("x", 100));
"#
        ),
        &["abababab", "-----", "100"]
    );
}

// ── preg_match_all collect matches ──────────────────────────────
#[test]
fn preg_match_all_collect() {
    assert_eq!(
        run_prints(
            r#"<?php
$count = preg_match_all('/\d+/', "abc123def456ghi789", $matches);
echo $count;
echo implode(",", $matches[0]);
"#
        ),
        &["3", "123,456,789"]
    );
}

// ── preg_split with PREG_SPLIT_NO_EMPTY ─────────────────────────
#[test]
fn preg_split_no_empty_flag() {
    assert_eq!(
        run_prints(
            r#"<?php
$parts = preg_split('/[\s,]+/', "one  two,,three   four", -1, PREG_SPLIT_NO_EMPTY);
echo count($parts);
echo implode("|", $parts);
"#
        ),
        &["4", "one|two|three|four"]
    );
}

// ── preg_quote ───────────────────────────────────────────────────
#[test]
fn preg_quote_special_chars() {
    assert_eq!(
        run_prints(
            r#"<?php
$pattern = preg_quote("$1.00 (today)", "/");
$quoted = preg_match("/" . $pattern . "/", '$1.00 (today)');
echo $quoted;
"#
        ),
        &["1"]
    );
}

// ── preg_replace_callback ────────────────────────────────────────
#[test]
fn preg_replace_callback_counter() {
    assert_eq!(
        run_prints(
            r#"<?php
$i = 0;
$result = preg_replace_callback('/\d+/', function($m) use (&$i) {
    $i++;
    return $m[0] * 2;
}, "a1 b2 c3");
echo $result;
echo $i;
"#
        ),
        &["a2 b4 c6", "3"]
    );
}

// ── sscanf parsing ───────────────────────────────────────────────
#[test]
fn sscanf_parsing() {
    assert_eq!(
        run_prints(
            r#"<?php
$parsed = sscanf("Age: 25", "Age: %d");
echo $parsed[0];
$parsed2 = sscanf("2024-01-15", "%d-%d-%d");
echo $parsed2[0];
echo $parsed2[1];
echo $parsed2[2];
"#
        ),
        &["25", "2024", "1", "15"]
    );
}

// ── html_entity_decode ───────────────────────────────────────────
#[test]
fn html_entity_decode_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
echo html_entity_decode("&lt;p&gt;Hello &amp; World&lt;/p&gt;");
echo html_entity_decode("&copy; 2024 &trade;");
"#
        ),
        &["<p>Hello & World</p>", "© 2024 ™"]
    );
}

// ── nl2br in multiline context ───────────────────────────────────
#[test]
fn nl2br_multiline() {
    assert_eq!(
        run_prints(
            r#"<?php
$text = "line1\nline2\nline3";
$result = nl2br($text);
echo substr_count($result, "<br />");
"#
        ),
        &["2"]
    );
}
