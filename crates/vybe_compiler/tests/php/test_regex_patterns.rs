use super::helpers::run_prints;

// ── preg_match_all ───────────────────────────────────────────────
#[test]
fn preg_match_all_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$str = "The price is $10 and $20 and $30";
$count = preg_match_all('/\$(\d+)/', $str, $matches);
echo $count;
echo implode(",", $matches[1]);
"#
        ),
        &["310,20,30"]
    );
}

#[test]
fn preg_match_all_emails() {
    assert_eq!(
        run_prints(
            r#"<?php
$text = "Contact alice@example.com or bob@test.org for info";
preg_match_all('/[\w.]+@[\w.]+/', $text, $matches);
echo implode(",", $matches[0]);
"#
        ),
        &["alice@example.com,bob@test.org"]
    );
}

// ── preg_replace_callback ────────────────────────────────────────
#[test]
fn preg_replace_callback_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$result = preg_replace_callback('/\d+/', function($matches) {
    return $matches[0] * 2;
}, "I have 5 apples and 3 oranges");
echo $result;
"#
        ),
        &["I have 10 apples and 6 oranges"]
    );
}

#[test]
fn preg_replace_callback_ucfirst() {
    assert_eq!(
        run_prints(
            r#"<?php
$result = preg_replace_callback('/\b\w/', function($m) {
    return strtoupper($m[0]);
}, "hello beautiful world");
echo $result;
"#
        ),
        &["Hello Beautiful World"]
    );
}

// ── preg_quote ───────────────────────────────────────────────────
#[test]
fn preg_quote_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$special = "price is $10.00 (USD)";
$escaped = preg_quote($special, '/');
echo preg_match('/' . $escaped . '/', $special) ? "found" : "not found";
"#
        ),
        &["found"]
    );
}

// ── Named groups ─────────────────────────────────────────────────
#[test]
fn named_groups_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$date = "2024-06-15";
preg_match('/(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})/', $date, $m);
echo $m['year'];
echo $m['month'];
echo $m['day'];
"#
        ),
        &["20240615"]
    );
}

#[test]
fn named_groups_all() {
    assert_eq!(
        run_prints(
            r#"<?php
$log = "ERROR: file not found\nWARN: low memory\nERROR: timeout";
preg_match_all('/(?P<level>ERROR|WARN): (?P<msg>.+)/', $log, $matches);
echo implode(",", $matches['level']);
echo implode(",", $matches['msg']);
"#
        ),
        &["ERROR,WARN,ERRORfile not found,low memory,timeout"]
    );
}

// ── Lookahead / lookbehind ───────────────────────────────────────
#[test]
fn lookahead_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$str = "100px 200em 300px 400rem";
preg_match_all('/\d+(?=px)/', $str, $matches);
echo implode(",", $matches[0]);
"#
        ),
        &["100,300"]
    );
}

#[test]
fn lookbehind_basic() {
    assert_eq!(
        run_prints(
            r#"<?php
$str = "$100 €200 $300 €400";
preg_match_all('/(?<=\$)\d+/', $str, $matches);
echo implode(",", $matches[0]);
"#
        ),
        &["100,300"]
    );
}

#[test]
fn negative_lookahead() {
    assert_eq!(
        run_prints(
            r#"<?php
$str = "foo123 bar456 foo789 baz000";
preg_match_all('/\b(?!foo)\w+\d+/', $str, $matches);
echo implode(",", $matches[0]);
"#
        ),
        &["bar456,baz000"]
    );
}

// ── preg_replace with backrefs ───────────────────────────────────
#[test]
fn preg_replace_backref() {
    assert_eq!(
        run_prints(
            r#"<?php
$str = "John Smith";
echo preg_replace('/(\w+) (\w+)/', '$2, $1', $str);
"#
        ),
        &["Smith, John"]
    );
}

#[test]
fn preg_replace_multiple_patterns() {
    assert_eq!(
        run_prints(
            r#"<?php
$str = "Hello   World   PHP";
$result = preg_replace('/\s+/', ' ', $str);
echo $result;
"#
        ),
        &["Hello World PHP"]
    );
}

// ── Character classes ────────────────────────────────────────────
#[test]
fn char_class_patterns() {
    assert_eq!(
        run_prints(
            r#"<?php
echo preg_match('/^\d+$/', "12345") ? "digits" : "no";
echo preg_match('/^[a-zA-Z]+$/', "Hello") ? "alpha" : "no";
echo preg_match('/^[\w]+$/', "hello_123") ? "word" : "no";
echo preg_match('/^[^aeiou]+$/i', "fly") ? "no vowels" : "has vowels";
"#
        ),
        &["digitsalphawordno vowels"]
    );
}

// ── Quantifiers ──────────────────────────────────────────────────
#[test]
fn quantifier_patterns() {
    assert_eq!(
        run_prints(
            r#"<?php
echo preg_match('/^a{3}$/', "aaa") ? "yes" : "no";
echo preg_match('/^a{2,4}$/', "aaa") ? "yes" : "no";
echo preg_match('/^a{2,4}$/', "aaaaa") ? "yes" : "no";
echo preg_match('/^a+$/', "aaa") ? "yes" : "no";
echo preg_match('/^a*$/', "") ? "yes" : "no";
echo preg_match('/^a?$/', "a") ? "yes" : "no";
"#
        ),
        &["yesyesnoyesyesyes"]
    );
}

// ── Real-world patterns ──────────────────────────────────────────
#[test]
fn validate_email_pattern() {
    assert_eq!(
        run_prints(
            r#"<?php
$pattern = '/^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/';
echo preg_match($pattern, "user@example.com") ? "valid" : "invalid";
echo preg_match($pattern, "bad@") ? "valid" : "invalid";
echo preg_match($pattern, "test.user+tag@domain.co.uk") ? "valid" : "invalid";
"#
        ),
        &["validinvalidvalid"]
    );
}

#[test]
fn extract_urls() {
    assert_eq!(
        run_prints(
            r#"<?php
$text = "Visit https://example.com and http://test.org/page for more";
preg_match_all('/https?:\/\/[\w.\/]+/', $text, $matches);
echo implode("\n", $matches[0]);
"#
        ),
        &["https://example.com\nhttp://test.org/page"]
    );
}

#[test]
fn slug_generation() {
    assert_eq!(
        run_prints(
            r#"<?php
function slugify(string $text): string {
    $text = strtolower($text);
    $text = preg_replace('/[^a-z0-9]+/', '-', $text);
    $text = trim($text, '-');
    return $text;
}
echo slugify("Hello World! This is PHP");
echo slugify("  Multiple   Spaces  Here  ");
"#
        ),
        &["hello-world-this-is-phpmultiple-spaces-here"]
    );
}

#[test]
fn html_tag_extraction() {
    assert_eq!(
        run_prints(
            r#"<?php
$html = '<div class="main"><p>Hello</p><p>World</p></div>';
preg_match_all('/<p>(.*?)<\/p>/', $html, $matches);
echo implode(",", $matches[1]);
"#
        ),
        &["Hello,World"]
    );
}

// ── preg_split ───────────────────────────────────────────────────
#[test]
fn preg_split_delimiters() {
    assert_eq!(
        run_prints(
            r#"<?php
$str = "one, two;three  four";
$parts = preg_split('/[\s,;]+/', $str);
echo implode("|", $parts);
"#
        ),
        &["one|two|three|four"]
    );
}

#[test]
fn preg_split_with_limit() {
    assert_eq!(
        run_prints(
            r#"<?php
$str = "a-b-c-d-e";
$parts = preg_split('/-/', $str, 3);
echo implode("|", $parts);
"#
        ),
        &["a|b|c-d-e"]
    );
}
