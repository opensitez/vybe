use super::helpers::run_prints;

// ── preg_replace with limit and count ────────────────────────

#[test] fn preg_replace_limit_one() {
    assert_eq!(run_prints(r#"<?php echo preg_replace('/a/', 'X', 'banana', 1); "#), vec!["bXnana"]);
}
#[test] fn preg_replace_count_param() {
    assert_eq!(run_prints(r#"<?php $c = 0; preg_replace('/a/', 'X', 'banana', -1, $c); echo $c; "#), vec!["3"]);
}
#[test] fn preg_replace_callback_transform() {
    assert_eq!(run_prints(r#"<?php
$r = preg_replace_callback('/\d+/', fn($m) => $m[0] * 2, 'I have 3 apples and 5 bananas');
echo $r;
"#), vec!["I have 6 apples and 10 bananas"]);
}
#[test] fn preg_replace_callback_array_patterns() {
    assert_eq!(run_prints(r#"<?php
$r = preg_replace_callback_array([
    '/\b[A-Z][a-z]+/' => fn($m) => strtolower($m[0]),
    '/\b\d+/' => fn($m) => $m[0] * 10,
], 'Hello 5 World 3');
echo $r;
"#), vec!["hello 50 world 30"]);
}

// ── preg_match_all ────────────────────────────────────────────

#[test] fn preg_match_all_returns_count() {
    assert_eq!(run_prints(r#"<?php echo preg_match_all('/\d+/', 'abc123def456ghi789'); "#), vec!["3"]);
}
#[test] fn preg_match_all_capture_groups() {
    assert_eq!(run_prints(r#"<?php
preg_match_all('/(\w+)=(\w+)/', 'a=1 b=2 c=3', $m);
echo implode(',', $m[1]) . ':' . implode(',', $m[2]);
"#), vec!["a,b,c:1,2,3"]);
}
#[test] fn preg_match_all_set_order() {
    assert_eq!(run_prints(r#"<?php
preg_match_all('/(\w+):(\d+)/', 'foo:1 bar:2', $m, PREG_SET_ORDER);
echo $m[0][1] . '=' . $m[0][2] . ',' . $m[1][1] . '=' . $m[1][2];
"#), vec!["foo=1,bar=2"]);
}

// ── preg_split ────────────────────────────────────────────────

#[test] fn preg_split_on_whitespace() {
    assert_eq!(run_prints(r#"<?php echo implode('-', preg_split('/\s+/', 'one  two   three')); "#), vec!["one-two-three"]);
}
#[test] fn preg_split_limit() {
    assert_eq!(run_prints(r#"<?php echo implode('|', preg_split('/:/', 'a:b:c:d', 3)); "#), vec!["a|b|c:d"]);
}
#[test] fn preg_split_no_empty() {
    assert_eq!(run_prints(r#"<?php echo count(preg_split('/,/', 'a,,b,,c', -1, PREG_SPLIT_NO_EMPTY)); "#), vec!["3"]);
}

// ── preg_grep ─────────────────────────────────────────────────

#[test] fn preg_grep_filters_matching() {
    assert_eq!(run_prints(r#"<?php
$a = ['foo','bar123','baz','qux456'];
echo implode(',', preg_grep('/\d/', $a));
"#), vec!["bar123,qux456"]);
}
#[test] fn preg_grep_invert() {
    assert_eq!(run_prints(r#"<?php
$a = ['foo','bar123','baz'];
echo implode(',', preg_grep('/\d/', $a, PREG_GREP_INVERT));
"#), vec!["foo,baz"]);
}

// ── Regex flags ───────────────────────────────────────────────

#[test] fn regex_case_insensitive() {
    assert_eq!(run_prints(r#"<?php echo preg_match('/HELLO/i', 'Hello World') ? 'yes' : 'no'; "#), vec!["yes"]);
}
#[test] fn regex_multiline_anchor() {
    assert_eq!(run_prints(r#"<?php
$n = preg_match_all('/^\d+/m', "1 foo\n2 bar\n3 baz");
echo $n;
"#), vec!["3"]);
}
#[test] fn regex_dotall_matches_newline() {
    assert_eq!(run_prints(r#"<?php echo preg_match('/a.b/s', "a\nb") ? 'yes' : 'no'; "#), vec!["yes"]);
}
#[test] fn regex_extended_ignores_whitespace() {
    assert_eq!(run_prints(r#"<?php echo preg_match('/\d +  \d/x', '1 2') ? 'yes' : 'no'; "#), vec!["yes"]);
}

// ── Backreferences ────────────────────────────────────────────

#[test] fn backreference_in_pattern() {
    assert_eq!(run_prints(r#"<?php echo preg_match('/(\w)\1/', 'hello') ? 'yes' : 'no'; "#), vec!["yes"]);
}
#[test] fn named_backreference_in_replace() {
    assert_eq!(run_prints(r#"<?php echo preg_replace('/(?P<word>\w+) \1/', '$1', 'hello hello world'); "#), vec!["hello world"]);
}

// ── Lookahead / lookbehind ────────────────────────────────────

#[test] fn positive_lookahead() {
    assert_eq!(run_prints(r#"<?php preg_match('/\d+(?= dollars)/', '100 dollars', $m); echo $m[0]; "#), vec!["100"]);
}
#[test] fn negative_lookahead() {
    assert_eq!(run_prints(r#"<?php echo preg_match('/\d+(?! dollars)/', '100 euros', $m) ? $m[0] : 'no'; "#), vec!["100"]);
}
#[test] fn lookbehind_positive() {
    assert_eq!(run_prints(r#"<?php preg_match('/(?<=USD )\d+/', 'USD 500', $m); echo $m[0]; "#), vec!["500"]);
}

// ── str_replace with array ────────────────────────────────────

#[test] fn str_replace_array_search() {
    assert_eq!(run_prints(r#"<?php echo str_replace(['a','e','i','o','u'], '*', 'hello world'); "#), vec!["h*ll* w*rld"]);
}
#[test] fn str_replace_array_pairs() {
    assert_eq!(run_prints(r#"<?php echo str_replace(['PHP','world'], ['Vybe','PHP'], 'PHP world'); "#), vec!["Vybe PHP"]);
}
#[test] fn str_ireplace_case_insensitive() {
    assert_eq!(run_prints(r#"<?php echo str_ireplace('HELLO', 'Hi', 'Hello World hello'); "#), vec!["Hi World Hi"]);
}
