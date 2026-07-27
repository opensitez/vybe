use super::helpers::run_prints;

// ── preg_replace with limit and count ────────────────────────

#[test]
fn preg_replace_limit_one() {
    assert_eq!(
        run_prints(r#"<?php echo preg_replace('/a/', 'X', 'banana', 1); "#),
        vec!["bXnana"]
    );
}
#[test]
fn preg_replace_count_param() {
    assert_eq!(
        run_prints(r#"<?php $c = 0; preg_replace('/a/', 'X', 'banana', -1, $c); echo $c; "#),
        vec!["3"]
    );
}
#[test]
fn preg_replace_callback_transform() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = preg_replace_callback('/\d+/', fn($m) => $m[0] * 2, 'I have 3 apples and 5 bananas');
echo $r;
echo "\n";
"#
        ),
        vec!["I have 6 apples and 10 bananas"]
    );
}
#[test]
fn preg_replace_callback_array_patterns() {
    assert_eq!(
        run_prints(
            r#"<?php
$r = preg_replace_callback_array([
    '/\b[A-Z][a-z]+/' => fn($m) => strtolower($m[0]),
    '/\b\d+/' => fn($m) => $m[0] * 10,
], 'Hello 5 World 3');
echo $r;
echo "\n";
"#
        ),
        vec!["hello 50 world 30"]
    );
}

// ── preg_match_all ────────────────────────────────────────────

#[test]
fn preg_match_all_returns_count() {
    assert_eq!(
        run_prints(r#"<?php echo preg_match_all('/\d+/', 'abc123def456ghi789'); "#),
        vec!["3"]
    );
}
#[test]
fn preg_match_all_capture_groups() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match_all('/(\w+)=(\w+)/', 'a=1 b=2 c=3', $m);
echo implode(',', $m[1]) . ':' . implode(',', $m[2]);
echo "\n";
"#
        ),
        vec!["a,b,c:1,2,3"]
    );
}
#[test]
fn preg_match_all_set_order() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match_all('/(\w+):(\d+)/', 'foo:1 bar:2', $m, PREG_SET_ORDER);
echo $m[0][1] . '=' . $m[0][2] . ',' . $m[1][1] . '=' . $m[1][2];
echo "\n";
"#
        ),
        vec!["foo=1,bar=2"]
    );
}

// ── preg_split ────────────────────────────────────────────────

#[test]
fn preg_split_on_whitespace() {
    assert_eq!(
        run_prints(r#"<?php echo implode('-', preg_split('/\s+/', 'one  two   three')); "#),
        vec!["one-two-three"]
    );
}
#[test]
fn preg_split_limit() {
    assert_eq!(
        run_prints(r#"<?php echo implode('|', preg_split('/:/', 'a:b:c:d', 3)); "#),
        vec!["a|b|c:d"]
    );
}
#[test]
fn preg_split_no_empty() {
    assert_eq!(
        run_prints(r#"<?php echo count(preg_split('/,/', 'a,,b,,c', -1, PREG_SPLIT_NO_EMPTY)); "#),
        vec!["3"]
    );
}

// ── preg_grep ─────────────────────────────────────────────────

#[test]
fn preg_grep_filters_matching() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['foo','bar123','baz','qux456'];
echo implode(',', preg_grep('/\d/', $a));
echo "\n";
"#
        ),
        vec!["bar123,qux456"]
    );
}
#[test]
fn preg_grep_invert() {
    assert_eq!(
        run_prints(
            r#"<?php
$a = ['foo','bar123','baz'];
echo implode(',', preg_grep('/\d/', $a, PREG_GREP_INVERT));
echo "\n";
"#
        ),
        vec!["foo,baz"]
    );
}

// ── Regex flags ───────────────────────────────────────────────

#[test]
fn regex_case_insensitive() {
    assert_eq!(
        run_prints(r#"<?php echo preg_match('/HELLO/i', 'Hello World') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn regex_multiline_anchor() {
    assert_eq!(
        run_prints(
            r#"<?php
$n = preg_match_all('/^\d+/m', "1 foo\n2 bar\n3 baz");
echo $n;
echo "\n";
"#
        ),
        vec!["3"]
    );
}
#[test]
fn regex_dotall_matches_newline() {
    assert_eq!(
        run_prints(r#"<?php echo preg_match('/a.b/s', "a\nb") ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn regex_extended_ignores_whitespace() {
    assert_eq!(
        // In /x (extended) mode unescaped whitespace is ignored, so the
        // pattern collapses to `\d+\d`, which cannot match "1 2" (space
        // between the digits). PHP 8.4 returns 0 here — verified.
        run_prints(r#"<?php echo preg_match('/\d +  \d/x', '1 2') ? 'yes' : 'no'; "#),
        vec!["no"]
    );
}

// ── Backreferences ────────────────────────────────────────────

#[test]
fn backreference_in_pattern() {
    assert_eq!(
        run_prints(r#"<?php echo preg_match('/(\w)\1/', 'hello') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn named_backreference_in_replace() {
    assert_eq!(
        run_prints(r#"<?php echo preg_replace('/(?P<word>\w+) \1/', '$1', 'hello hello world'); "#),
        vec!["hello world"]
    );
}

// ── Lookahead / lookbehind ────────────────────────────────────

#[test]
fn positive_lookahead() {
    assert_eq!(
        run_prints(r#"<?php preg_match('/\d+(?= dollars)/', '100 dollars', $m); echo $m[0]; "#),
        vec!["100"]
    );
}
#[test]
fn negative_lookahead() {
    assert_eq!(
        run_prints(
            r#"<?php echo preg_match('/\d+(?! dollars)/', '100 euros', $m) ? $m[0] : 'no'; "#
        ),
        vec!["100"]
    );
}
#[test]
fn lookbehind_positive() {
    assert_eq!(
        run_prints(r#"<?php preg_match('/(?<=USD )\d+/', 'USD 500', $m); echo $m[0]; "#),
        vec!["500"]
    );
}

// ── str_replace with array ────────────────────────────────────

#[test]
fn str_replace_array_search() {
    assert_eq!(
        run_prints(r#"<?php echo str_replace(['a','e','i','o','u'], '*', 'hello world'); "#),
        vec!["h*ll* w*rld"]
    );
}
#[test]
fn str_replace_array_pairs() {
    assert_eq!(
        run_prints(r#"<?php echo str_replace(['PHP','world'], ['Vybe','PHP'], 'PHP world'); "#),
        vec!["Vybe PHP"]
    );
}
#[test]
fn str_ireplace_case_insensitive() {
    assert_eq!(
        run_prints(r#"<?php echo str_ireplace('HELLO', 'Hi', 'Hello World hello'); "#),
        vec!["Hi World Hi"]
    );
}

#[test]
fn preg_quote_escapes_delimiter_and_meta() {
    assert_eq!(
        run_prints(
            r#"<?php $pat = '/'.preg_quote('/foo.bar', '/').'/'; echo preg_match($pat, 'a/foo.bar/'); "#
        ),
        vec!["1"]
    );
}

#[test]
fn preg_match_with_optional_group() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match('/^(?<x>\w+)(?:\s+(\d+))?$/', 'name 123', $m);
echo isset($m[2]) ? $m[2] : 'none';
echo '|';
preg_match('/^(?<x>\w+)(?:\s+(\d+))?$/', 'name', $m2);
echo isset($m2[2]) ? $m2[2] : 'none';
"#
        ),
        vec!["123|none"]
    );
}

#[test]
fn preg_match_all_with_no_match() {
    assert_eq!(
        run_prints(
            r#"<?php echo preg_match_all('/\d{4}/', 'no-digits-here', $m); echo '|'; echo count($m[0]); "#
        ),
        vec!["0|0"]
    );
}

#[test]
fn regex_offset_capture_starts_after_offset() {
    assert_eq!(
        run_prints(
            r#"<?php
$matches = [];
$result = preg_match('/a/', 'abcab', $matches, PREG_OFFSET_CAPTURE, 2);
echo $result;
echo '|';
echo $matches[0][0];
echo '|';
echo $matches[0][1];
"#
        ),
        vec!["1|a|3"]
    );
}

#[test]
fn regex_replace_backreference_reordering() {
    assert_eq!(
        run_prints(r#"<?php echo preg_replace('/(foo)(bar)/', '$2-$1', 'foobar'); "#),
        vec!["bar-foo"]
    );
}

#[test]
fn regex_unicode_property_class() {
    assert_eq!(
        run_prints(
            r#"<?php
echo preg_match_all('/\p{Lu}+/u', 'ABCé', $m);
echo '|';
echo $m[0][0] ?? 'none';
"#
        ),
        vec!["1|ABC"]
    );
}

#[test]
fn regex_split_offsets_and_no_captures() {
    assert_eq!(
        run_prints(
            r#"<?php
$parts = preg_split('/:/', 'a:b:c', -1, PREG_SPLIT_OFFSET_CAPTURE);
echo $parts[0][0];
echo '|';
echo $parts[1][0];
echo '|';
echo $parts[1][1];
"#
        ),
        vec!["a|b|2"]
    );
}

#[test]
fn preg_replace_backreference_with_grouped_digits_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$value = 'A1,B2,C3';
echo preg_replace('/([A-Z])(\d)/', '$2$1', $value);
"#
        ),
        vec!["1A,2B,3C"]
    );
}

#[test]
fn preg_match_offset_flag_only_after_match() {
    assert_eq!(
        run_prints(
            r#"<?php
$matches = [];
preg_match('/b/', 'abcab', $matches, PREG_OFFSET_CAPTURE);
echo $matches[0][1];
"#
        ),
        vec!["1"]
    );
}
