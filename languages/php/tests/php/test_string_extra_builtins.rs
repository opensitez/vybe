use super::helpers::{compile_ok, run_prints};

// ── addcslashes ──────────────────────────────────────────────────
#[test]
fn addcslashes_c_style_escapes() {
    compile_ok(
        r#"<?php
$s = "Hello\tWorld\n";
$escaped = addcslashes($s, "\t\n");
echo is_string($escaped) ? "ok" : "fail";
"#,
    );
}

#[test]
fn addcslashes_character_range_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo addcslashes("AZ", "A");
echo '|';
echo addcslashes("abc", "a..z");
"#
        ),
        vec!["\\A\\Z|\\a\\b\\c"]
    );
}

// ── stripcslashes ────────────────────────────────────────────────
#[test]
fn stripcslashes_remove_c_style_escapes() {
    compile_ok(
        r#"<?php
$escaped = 'He said \\"hello\\"';
$s = stripcslashes($escaped);
echo is_string($s) ? "ok" : "fail";
"#,
    );
}

#[test]
fn stripcslashes_revert_escapes_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo stripcslashes('tab\\tchar');
echo '|';
echo stripcslashes('back\\\\slash');
"#
        ),
        vec!["tab\tchar|back\\slash"]
    );
}

// ── htmlentities ─────────────────────────────────────────────────
#[test]
fn htmlentities_convert_all_applicable_chars() {
    compile_ok(
        r#"<?php
$html = '<a href="test">link & "quotes"</a>';
$encoded = htmlentities($html);
echo is_string($encoded) ? "ok" : "fail";
echo strpos($encoded, "&lt;") !== false ? "has-lt" : "no-lt";
"#,
    );
}

// ── html_entity_decode ───────────────────────────────────────────
#[test]
fn html_entity_decode_chars_from_entities() {
    compile_ok(
        r#"<?php
$encoded = "&lt;p&gt;Hello &amp; World&lt;/p&gt;";
$decoded = html_entity_decode($encoded);
echo strpos($decoded, "<p>") !== false ? "has-tag" : "no-tag";
echo strpos($decoded, "&") !== false ? "has-amp" : "no-amp";
"#,
    );
}

#[test]
fn htmlentities_and_decode_quote_modes_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$encoded = htmlentities('<a>"&\'', ENT_QUOTES);
echo str_contains($encoded, '&lt;') ? 'lt' : 'no';
echo '|';
echo str_contains($encoded, '&quot;') ? 'dq' : 'no';
echo '|';
echo str_contains($encoded, '&#039;') ? 'sq' : 'no';
echo '|';
echo html_entity_decode($encoded, ENT_QUOTES);
"#
        ),
        vec!["lt|dq|sq|<a>\"'"]
    );
}

// ── strcspn ──────────────────────────────────────────────────────
#[test]
fn strcspn_initial_segment_not_in_mask() {
    compile_ok(
        r#"<?php
$n = strcspn("abcdefg", "deh");
echo $n;
echo strcspn("hello", "aeiou");
echo strcspn("", "abc");
"#,
    );
}

// ── strpbrk ──────────────────────────────────────────────────────
#[test]
fn strpbrk_search_for_any_char_in_set() {
    compile_ok(
        r#"<?php
$result = strpbrk("This is a test", "aeiou");
echo is_string($result) ? "found" : "not";
$none = strpbrk("bcdfg", "aeiou");
echo $none === false ? "false" : "found";
"#,
    );
}

// ── str_shuffle ──────────────────────────────────────────────────
#[test]
fn str_shuffle_randomize_characters() {
    compile_ok(
        r#"<?php
$original = "abcdefghij";
$shuffled = str_shuffle($original);
echo strlen($shuffled) === strlen($original) ? "same-len" : "diff-len";
echo is_string($shuffled) ? "ok" : "fail";
"#,
    );
}

// ── strtok ───────────────────────────────────────────────────────
#[test]
fn strtok_tokenize_by_delimiter() {
    compile_ok(
        r#"<?php
$token = strtok("Hello World PHP", " ");
$parts = [];
while ($token !== false) {
    $parts[] = $token;
    $token = strtok(" ");
}
echo count($parts);
echo $parts[0];
"#,
    );
}

// ── substr_compare ───────────────────────────────────────────────
#[test]
fn substr_compare_binary_safe_from_offset() {
    compile_ok(
        r#"<?php
$result = substr_compare("abcdefg", "cde", 2, 3);
echo $result === 0 ? "equal" : "not-equal";
$diff = substr_compare("abcdefg", "xyz", 0, 3);
echo $diff !== 0 ? "different" : "same";
"#,
    );
}

// ── parse_str ────────────────────────────────────────────────────
#[test]
fn parse_str_query_string_into_variables() {
    compile_ok(
        r#"<?php
parse_str("name=Alice&age=30&city=Paris", $output);
echo $output["name"];
echo $output["age"];
echo $output["city"];
"#,
    );
}

#[test]
fn parse_str_nested_brackets_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
parse_str('user[name]=john&user[age]=32&tags[]=a&tags[]=b', $out);
echo $out['user']['name'];
echo '|';
echo $out['user']['age'];
echo '|';
echo $out['tags'][0] . ',' . $out['tags'][1];
"#
        ),
        vec!["john|32|a,b"]
    );
}

#[test]
fn str_shuffle_keeps_length_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$s = "abcdefgh";
$shuffled = str_shuffle($s);
echo strlen($s);
echo "|";
echo strlen($shuffled);
echo "|";
echo is_string($shuffled) ? "string" : "not";
"#
        ),
        vec!["8|8|string"]
    );
}

#[test]
fn stripcslashes_escapes_quotes_backslashes_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo stripcslashes("a\\nb\\\"c\\'d\\\\e");
echo "|";
echo stripcslashes("line1\\nline2");
"#
        ),
        vec!["a\nb\"c\\d|line1\nline2"]
    );
}

#[test]
fn preg_filter_returns_null_when_no_match_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$input = ["a", "b", "c"];
$result = preg_filter('/z/', 'X$0', $input);
echo var_export($result, true);
"#
        ),
        vec!["NULL"]
    );
}

// ── preg_filter ──────────────────────────────────────────────────
#[test]
fn preg_filter_return_only_matched_replaced() {
    compile_ok(
        r#"<?php
$input = ["foo1", "bar", "foo2", "baz", "foo3"];
$result = preg_filter('/^foo(\d)/', 'match$1', $input);
echo count($result);
echo is_array($result) ? "array" : "not";
"#,
    );
}

// ── preg_grep ────────────────────────────────────────────────────
#[test]
fn preg_grep_array_elements_matching_pattern() {
    compile_ok(
        r#"<?php
$numbers = [1, 15, 3, 200, 42, 7, 100];
$large = preg_grep('/^[0-9]{3}/', array_map('strval', $numbers));
echo is_array($large) ? "array" : "not";
"#,
    );
}

#[test]
fn preg_grep_with_integer_and_string_filter_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$values = ["a1", "b2", "a3", "c4"];
$matches = preg_grep('/^a/', $values);
echo count($matches);
echo '|';
echo implode(',', $matches);
"#
        ),
        vec!["2|a1,a3"]
    );
}

#[test]
fn preg_match_error_code_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match('/(/', 'abc');
echo preg_last_error() > 0 ? 'error' : 'ok';
"#
        ),
        vec!["error"]
    );
}

// ── preg_last_error ──────────────────────────────────────────────
#[test]
fn preg_last_error_after_match() {
    compile_ok(
        r#"<?php
preg_match('/\d+/', 'abc123');
$err = preg_last_error();
echo is_int($err) ? "int" : "not-int";
echo $err === PREG_NO_ERROR ? "no-error" : "error";
"#,
    );
}

// ── preg_match with named capture groups ────────────────────────
#[test]
fn preg_match_named_capture_groups() {
    compile_ok(
        r#"<?php
$date = "2024-07-15";
preg_match('/(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})/', $date, $m);
echo $m["year"];
echo $m["month"];
echo $m["day"];
"#,
    );
}

// ── preg_match_all collecting all groups ────────────────────────
#[test]
fn preg_match_all_collect_all_groups() {
    compile_ok(
        r#"<?php
$html = '<a href="http://example.com">one</a> <a href="http://test.org">two</a>';
$count = preg_match_all('/<a href="([^"]+)">([^<]+)<\/a>/', $html, $matches);
echo $count;
echo count($matches[1]);
echo $matches[2][0];
"#,
    );
}

// ── quoted_printable_encode ──────────────────────────────────────
#[test]
fn quoted_printable_encode_non_ascii() {
    compile_ok(
        r#"<?php
$text = "Subject line with special chars: \xc3\xa9\xc3\xa0";
$encoded = quoted_printable_encode($text);
echo is_string($encoded) ? "ok" : "fail";
"#,
    );
}

// ── quoted_printable_decode ──────────────────────────────────────
#[test]
fn quoted_printable_decode_encoded_input() {
    compile_ok(
        r#"<?php
$encoded = "Subject: =?UTF-8?Q?Hello=20World?=";
$decoded = quoted_printable_decode($encoded);
echo is_string($decoded) ? "ok" : "fail";
"#,
    );
}

// ── nl2br with xhtml parameter ──────────────────────────────────
#[test]
fn nl2br_with_xhtml_parameter() {
    compile_ok(
        r#"<?php
$text = "line one\nline two\nline three";
$xhtml = nl2br($text, true);
echo strpos($xhtml, "<br />") !== false ? "xhtml-br" : "not-found";
$html = nl2br($text, false);
echo strpos($html, "<br>") !== false ? "html-br" : "not-found";
"#,
    );
}

// ── number_format with custom separators ────────────────────────
#[test]
fn number_format_custom_decimal_and_thousands_separators() {
    compile_ok(
        r#"<?php
$n = 1234567.891;
echo number_format($n, 2, ',', '.');
echo number_format($n, 3, '/', '_');
echo number_format(0.5, 1, ',', '');
"#,
    );
}

// ── fprintf ──────────────────────────────────────────────────────
#[test]
fn fprintf_write_formatted_to_stream() {
    compile_ok(
        r#"<?php
$written = fprintf(STDOUT, "Name: %s, Age: %d, Score: %.2f\n", "Alice", 30, 98.5);
echo is_int($written) ? "int" : "not-int";
"#,
    );
}
