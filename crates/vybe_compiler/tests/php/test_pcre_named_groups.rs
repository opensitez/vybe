use super::helpers::{compile_ok, run_prints};

// ── Named capture groups (?P<name>...) ────────────────────────

#[test]
fn named_group_basic_match() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match('/(?P<year>\d{4})-(?P<month>\d{2})-(?P<day>\d{2})/', '2024-03-15', $m);
echo $m['year'] . ',' . $m['month'] . ',' . $m['day'];
"#
        ),
        vec!["2024,03,15"]
    );
}

#[test]
fn named_group_alternate_syntax() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match('/(?<first>\w+)\s+(?<last>\w+)/', 'John Doe', $m);
echo $m['first'] . ' ' . $m['last'];
"#
        ),
        vec!["John Doe"]
    );
}

#[test]
fn named_group_accessible_by_numeric_index_too() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match('/(?P<code>[A-Z]+)(\d+)/', 'ABC123', $m);
echo $m['code'] . ',' . $m[2];
"#
        ),
        vec!["ABC,123"]
    );
}

#[test]
fn named_group_in_replace_backreference() {
    assert_eq!(
        run_prints(
            r#"<?php
$result = preg_replace('/(?P<last>\w+),\s*(?P<first>\w+)/', '${first} ${last}', 'Smith, John');
echo $result;
"#
        ),
        vec!["John Smith"]
    );
}

// ── preg_match_all with named groups ─────────────────────────

#[test]
fn preg_match_all_named_groups_collect_all() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match_all('/(?P<word>[a-z]+)/', 'hello world foo', $m);
echo implode(',', $m['word']);
"#
        ),
        vec!["hello,world,foo"]
    );
}

#[test]
fn preg_match_all_named_groups_set_order() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match_all('/(?P<k>\w+)=(?P<v>\d+)/', 'a=1 b=2 c=3', $m, PREG_SET_ORDER);
$pairs = array_map(fn($e) => $e['k'] . ':' . $e['v'], $m);
echo implode(',', $pairs);
"#
        ),
        vec!["a:1,b:2,c:3"]
    );
}

// ── Lookahead assertions ──────────────────────────────────────

#[test]
fn positive_lookahead_matches_without_consuming() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match_all('/\d+(?= dollars)/', 'I have 50 dollars and 30 euros', $m);
echo implode(',', $m[0]);
"#
        ),
        vec!["50"]
    );
}

#[test]
fn negative_lookahead_excludes_followed_by() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match_all('/\d+(?! dollars)/', '50 dollars 30 euros', $m);
echo implode(',', $m[0]);
"#
        ),
        vec!["30"]
    );
}

#[test]
fn lookahead_in_split() {
    assert_eq!(
        run_prints(
            r#"<?php
$parts = preg_split('/(?=[A-Z])/', 'CamelCaseString');
$parts = array_filter($parts);
echo implode(',', array_values($parts));
"#
        ),
        vec!["Camel,Case,String"]
    );
}

// ── Lookbehind assertions ─────────────────────────────────────

#[test]
fn positive_lookbehind_matches_preceded_by() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match_all('/(?<=\$)\d+/', 'price $100 and $200', $m);
echo implode(',', $m[0]);
"#
        ),
        vec!["100,200"]
    );
}

#[test]
fn negative_lookbehind_excludes_preceded_by() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match_all('/(?<!\$)\b\d+\b/', 'price $100 and 200 items', $m);
echo implode(',', $m[0]);
"#
        ),
        vec!["200"]
    );
}

// ── preg_replace_callback_array ──────────────────────────────

#[test]
fn preg_replace_callback_array_multiple_patterns() {
    assert_eq!(
        run_prints(
            r#"<?php
$result = preg_replace_callback_array([
    '/\bfoo\b/' => fn($m) => 'bar',
    '/\bhello\b/' => fn($m) => 'goodbye',
], 'hello foo world');
echo $result;
"#
        ),
        vec!["goodbye bar world"]
    );
}

#[test]
fn preg_replace_callback_array_with_captures() {
    assert_eq!(
        run_prints(
            r#"<?php
$result = preg_replace_callback_array([
    '/(\d+) USD/' => fn($m) => $m[1] * 2 . ' USD',
    '/(\d+) EUR/' => fn($m) => $m[1] * 3 . ' EUR',
], '10 USD and 5 EUR');
echo $result;
"#
        ),
        vec!["20 USD and 15 EUR"]
    );
}

#[test]
fn preg_replace_callback_array_no_match_unchanged() {
    assert_eq!(
        run_prints(
            r#"<?php
$result = preg_replace_callback_array([
    '/xyz/' => fn($m) => 'replaced',
], 'no match here');
echo $result;
"#
        ),
        vec!["no match here"]
    );
}

// ── Non-capturing groups ──────────────────────────────────────

#[test]
fn non_capturing_group_not_in_matches() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match('/(?:foo|bar)(baz)/', 'foobaz', $m);
echo count($m) . ',' . $m[1];
"#
        ),
        vec!["2,baz"]
    );
}

// ── Atomic groups ─────────────────────────────────────────────

#[test]
fn possessive_quantifier_prevents_backtrack() {
    assert_eq!(
        run_prints(
            r#"<?php
$result = preg_match('/a++b/', 'aaab');
echo $result;
"#
        ),
        vec!["1"]
    );
}

// ── Multi-line and single-line flags ──────────────────────────

#[test]
fn multiline_flag_anchors_each_line() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match_all('/^\w+/m', "hello world\nfoo bar\nbaz", $m);
echo implode(',', $m[0]);
"#
        ),
        vec!["hello,foo,baz"]
    );
}

#[test]
fn dotall_flag_dot_matches_newlines() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match('/start(.+)end/s', "start\nline\nend", $m);
echo str_replace("\n", '|', $m[1]);
"#
        ),
        vec!["|line|"]
    );
}

// ── Extended mode with comments ───────────────────────────────

#[test]
fn extended_mode_ignores_whitespace() {
    assert_eq!(
        run_prints(
            r#"<?php
$pattern = '/
    (\d{4})  # year
    -
    (\d{2})  # month
/x';
preg_match($pattern, '2024-03', $m);
echo $m[1] . '-' . $m[2];
"#
        ),
        vec!["2024-03"]
    );
}

// ── Unicode support ───────────────────────────────────────────

#[test]
fn unicode_flag_matches_unicode_word_chars() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match_all('/\w+/u', 'café naïve', $m);
echo count($m[0]);
"#
        ),
        vec!["2"]
    );
}

// ── Backreference in pattern ──────────────────────────────────

#[test]
fn backreference_in_pattern_matches_repeated_word() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match('/(\w+) \1/', 'hello hello world', $m);
echo $m[1];
"#
        ),
        vec!["hello"]
    );
}

#[test]
fn named_backreference_k_syntax() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match('/(?P<tag>\w+).*\k<tag>/', 'div content div', $m);
echo $m['tag'];
"#
        ),
        vec!["div"]
    );
}

// ── preg_quote edge cases ─────────────────────────────────────

#[test]
fn preg_quote_escapes_special_chars() {
    assert_eq!(
        run_prints(
            r#"<?php
$pattern = preg_quote('.+*?()', '/');
echo preg_match('/' . $pattern . '/', '.+*?()') ? 'yes' : 'no';
"#
        ),
        vec!["yes"]
    );
}

// ── preg_split with limit ─────────────────────────────────────

#[test]
fn preg_split_with_limit() {
    assert_eq!(
        run_prints(
            r#"<?php
$parts = preg_split('/,/', 'a,b,c,d,e', 3);
echo implode('|', $parts);
"#
        ),
        vec!["a|b|c,d,e"]
    );
}

// ── preg_grep filters array by pattern ───────────────────────

#[test]
fn preg_grep_filters_matching_elements() {
    assert_eq!(
        run_prints(
            r#"<?php
$arr = ['apple', 'banana', 'apricot', 'cherry'];
$ap = preg_grep('/^ap/', $arr);
echo implode(',', $ap);
"#
        ),
        vec!["apple,apricot"]
    );
}

#[test]
fn preg_grep_inverted_filter() {
    assert_eq!(
        run_prints(
            r#"<?php
$nums = ['1', '2a', '3', '4b', '5'];
$notPure = preg_grep('/^\d+$/', $nums, PREG_GREP_INVERT);
echo implode(',', $notPure);
"#
        ),
        vec!["2a,4b"]
    );
}
