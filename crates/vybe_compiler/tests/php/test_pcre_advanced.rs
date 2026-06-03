use super::helpers::compile_ok;

// ── Named capture groups (?P<name>...) ───────────────────────────

#[test]
fn named_capture_groups() {
    compile_ok(
        r#"<?php
$date = '2024-06-15';
preg_match('/(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})/', $date, $m);
echo $m['year'] . '-' . $m['month'] . '-' . $m['day'];
"#,
    );
}

// ── Non-capturing group (?:...) ──────────────────────────────────

#[test]
fn non_capturing_group() {
    compile_ok(
        r#"<?php
preg_match('/(?:foo|bar)(baz)/', 'foobaz', $m);
echo $m[0];
echo $m[1];
echo isset($m[2]) ? 'has 2' : 'no 2';
"#,
    );
}

// ── Lookahead (?=...) ────────────────────────────────────────────

#[test]
fn positive_lookahead() {
    compile_ok(
        r#"<?php
preg_match_all('/\d+(?= dollars)/', 'I have 100 dollars and 50 cents', $m);
echo implode(',', $m[0]);
"#,
    );
}

// ── Negative lookahead (?!...) ───────────────────────────────────

#[test]
fn negative_lookahead() {
    compile_ok(
        r#"<?php
preg_match_all('/\bfoo(?!bar)\w*/', 'foobar foobaz fooqwe', $m);
echo implode(',', $m[0]);
"#,
    );
}

// ── Lookbehind (?<=...) ──────────────────────────────────────────

#[test]
fn positive_lookbehind() {
    compile_ok(
        r#"<?php
preg_match_all('/(?<=USD )\d+/', 'USD 100 EUR 200 USD 300', $m);
echo implode(',', $m[0]);
"#,
    );
}

// ── Negative lookbehind (?<!...) ─────────────────────────────────

#[test]
fn negative_lookbehind() {
    compile_ok(
        r#"<?php
preg_match_all('/(?<!USD )\b\d+/', 'USD 100 EUR 200 CAD 300', $m);
echo implode(',', $m[0]);
"#,
    );
}

// ── Non-greedy quantifier .*? ─────────────────────────────────────

#[test]
fn non_greedy_quantifier() {
    compile_ok(
        r#"<?php
$html = '<b>bold</b> and <b>more</b>';
preg_match_all('/<b>.*?<\/b>/', $html, $m);
echo count($m[0]);
echo implode(',', $m[0]);
"#,
    );
}

// ── Possessive quantifier ─────────────────────────────────────────

#[test]
fn possessive_quantifier() {
    compile_ok(
        r#"<?php
// PHP PCRE supports possessive via ++, *+, ?+
$pattern = '/^\w++$/';
echo preg_match($pattern, 'hello123') ? 'matched' : 'no match';
echo preg_match($pattern, 'hello world') ? 'matched' : 'no match';
"#,
    );
}

// ── Atomic group (?>...) ─────────────────────────────────────────

#[test]
fn atomic_group() {
    compile_ok(
        r#"<?php
// Atomic group prevents backtracking into it
$pattern = '/(?>a|ab)c/';
echo preg_match($pattern, 'abc') ? 'matched' : 'no match';
echo preg_match($pattern, 'ac') ? 'matched' : 'no match';
"#,
    );
}

// ── Unicode mode /u flag ─────────────────────────────────────────

#[test]
fn unicode_mode_flag() {
    compile_ok(
        r#"<?php
$str = 'Héllo Wörld';
preg_match_all('/\p{L}+/u', $str, $m);
echo implode(' ', $m[0]);
echo preg_match('/^\p{Lu}/u', 'Ñoño') ? ':uppercase start' : ':no uppercase start';
"#,
    );
}

// ── Multiline mode /m flag ───────────────────────────────────────

#[test]
fn multiline_mode_flag() {
    compile_ok(
        r#"<?php
$text = "first line\nsecond line\nthird line";
preg_match_all('/^\w+/m', $text, $m);
echo implode(',', $m[0]);
"#,
    );
}

// ── Dotall /s mode (dot matches newline) ─────────────────────────

#[test]
fn dotall_s_flag() {
    compile_ok(
        r#"<?php
$text = "start\nmiddle\nend";
echo preg_match('/start.+end/', $text) ? 'matched' : 'no match';
echo preg_match('/start.+end/s', $text) ? 'matched' : 'no match';
"#,
    );
}

// ── Extended /x mode (whitespace ignored) ───────────────────────

#[test]
fn extended_x_flag() {
    compile_ok(
        r#"<?php
$pattern = '/
    ^           # start
    \d{4}       # year
    -           # separator
    \d{2}       # month
    -           # separator
    \d{2}       # day
    $           # end
/x';
echo preg_match($pattern, '2024-06-15') ? 'matched' : 'no match';
echo preg_match($pattern, '24-6-1') ? 'matched' : 'no match';
"#,
    );
}

// ── Case-insensitive /i with Unicode ─────────────────────────────

#[test]
fn case_insensitive_with_unicode() {
    compile_ok(
        r#"<?php
echo preg_match('/héllo/iu', 'HÉLLO') ? 'matched' : 'no match';
echo preg_match('/\p{Lu}+/u', 'ABC') ? 'matched' : 'no match';
"#,
    );
}

// ── preg_match_all with PREG_SET_ORDER ───────────────────────────

#[test]
fn preg_match_all_set_order() {
    compile_ok(
        r#"<?php
preg_match_all('/(\d{4})-(\d{2})/', '2024-01 and 2024-06', $m, PREG_SET_ORDER);
echo count($m);
echo $m[0][0] . ':' . $m[0][1] . ':' . $m[0][2];
echo $m[1][0] . ':' . $m[1][1] . ':' . $m[1][2];
"#,
    );
}

// ── preg_match_all with PREG_OFFSET_CAPTURE ──────────────────────

#[test]
fn preg_match_all_offset_capture() {
    compile_ok(
        r#"<?php
preg_match_all('/\d+/', 'abc123def456', $m, PREG_OFFSET_CAPTURE);
echo $m[0][0][0] . '@' . $m[0][0][1];
echo $m[0][1][0] . '@' . $m[0][1][1];
"#,
    );
}

// ── preg_replace with backreference ─────────────────────────────

#[test]
fn preg_replace_backreference() {
    compile_ok(
        r#"<?php
$result = preg_replace('/(\w+)\s+(\w+)/', '$2 $1', 'Hello World');
echo $result;
$result2 = preg_replace('/(\d{4})-(\d{2})-(\d{2})/', '$3/$2/$1', '2024-06-15');
echo $result2;
"#,
    );
}

// ── preg_replace_callback returning modified match ───────────────

#[test]
fn preg_replace_callback_modify_match() {
    compile_ok(
        r#"<?php
$result = preg_replace_callback('/\b(\w)(\w+)\b/', function($m) {
    return strtoupper($m[1]) . strtolower($m[2]);
}, 'hello world from php');
echo $result;
"#,
    );
}

// ── preg_split with PREG_SPLIT_DELIM_CAPTURE ─────────────────────

#[test]
fn preg_split_delim_capture() {
    compile_ok(
        r#"<?php
$parts = preg_split('/([\s,;]+)/', 'one, two; three four', -1, PREG_SPLIT_DELIM_CAPTURE);
// Result includes delimiters as captured groups
echo count($parts) > 4 ? 'has delimiters' : 'no delimiters';
echo $parts[0];
"#,
    );
}

// ── preg_quote escaping special characters ───────────────────────

#[test]
fn preg_quote_special_chars() {
    compile_ok(
        r#"<?php
$user_input = 'price is $10.00 (USD) + tax?';
$escaped = preg_quote($user_input, '/');
echo preg_match('/' . $escaped . '/', $user_input) ? 'found' : 'not found';
$special = '\.+*?[^]$(){}=!<>|:-#';
$q = preg_quote($special);
echo is_string($q) ? 'quoted' : 'fail';
"#,
    );
}
