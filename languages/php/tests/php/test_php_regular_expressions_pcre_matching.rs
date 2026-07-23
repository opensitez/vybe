use super::helpers::{compile_ok, run_prints};

// ═══════════════════════════════════════════════════════════
// PHP: Regular Expressions & PCRE — preg_match, preg_match_all, preg_replace, preg_replace_callback, preg_split, named groups
// ═══════════════════════════════════════════════════════════

#[test]
fn test_php_preg_match_named_capturing_groups() {
    let out = run_prints(
        r#"<?php
$pattern = '/(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})/';
$subject = "Date: 2024-05-12";
if (preg_match($pattern, $subject, $matches)) {
    echo "Year={$matches['year']} Month={$matches['month']} Day={$matches['day']}";
}
"#,
    );
    assert_eq!(out, vec!["Year=2024 Month=05 Day=12"]);
}

#[test]
fn test_php_preg_replace_callback_callable() {
    let out = run_prints(
        r#"<?php
$input = "word1 word2 word3";
$result = preg_replace_callback('/\b\w+\b/', function($matches) {
    return strtoupper($matches[0]);
}, $input);
echo $result;
"#,
    );
    assert_eq!(out, vec!["WORD1 WORD2 WORD3"]);
}

#[test]
fn test_php_preg_match_all_global_extraction() {
    let out = run_prints(
        r#"<?php
$text = "Emails: alice@domain.com, bob@example.org";
preg_match_all('/[\w.-]+@[\w.-]+/', $text, $matches);
echo implode(", ", $matches[0]);
"#,
    );
    assert_eq!(out, vec!["alice@domain.com, bob@example.org"]);
}

#[test]
fn test_php_preg_split_pattern_splitting() {
    let out = run_prints(
        r#"<?php
$keywords = preg_split('/[\s,]+/', "hypertext language, programming");
echo implode("|", $keywords);
"#,
    );
    assert_eq!(out, vec!["hypertext|language|programming"]);
}

#[test]
fn test_php_preg_quote_regex_escaping() {
    let out = run_prints(
        r#"<?php
$rawText = "price is $10.00 (tax incl.)";
$escaped = preg_quote($rawText, '/');
echo preg_match('/' . $escaped . '/', $rawText) ? "MATCHED" : "NO_MATCH";
"#,
    );
    assert_eq!(out, vec!["MATCHED"]);
}

#[test]
fn test_php_preg_replace_array_patterns_replacements() {
    compile_ok(
        r#"<?php
$patterns = ['/quick/', '/brown/', '/fox/'];
$replacements = ['bear', 'black', 'wolf'];
echo preg_replace($patterns, $replacements, 'The quick brown fox jumps');
"#,
    );
}

#[test]
fn test_php_preg_last_error_and_msg() {
    compile_ok(
        r#"<?php
@preg_match('/(?:\D+)+/', '12345678901234567890');
if (preg_last_error() !== PREG_NO_ERROR) {
    echo "PCRE Error: " . preg_last_error_msg();
}
"#,
    );
}

#[test]
fn test_php_preg_match_flags_offset_capture() {
    compile_ok(
        r#"<?php
$str = "abc 123 def";
preg_match('/\d+/', $str, $matches, PREG_OFFSET_CAPTURE);
echo "Val={$matches[0][0]} Offset={$matches[0][1]}";
"#,
    );
}

#[test]
fn test_php_preg_replace_count_limit() {
    compile_ok(
        r#"<?php
$count = 0;
$result = preg_replace('/foo/', 'bar', 'foo foo foo', limit: 2, count: $count);
echo "Result=$result Count=$count";
"#,
    );
}

#[test]
fn test_php_preg_grep_array_filtering() {
    compile_ok(
        r#"<?php
$input = ["123", "abc", "456", "def"];
$numbers = preg_grep('/^\d+$/', $input);
echo implode(",", $numbers);
"#,
    );
}
