//! `preg_match`, `preg_replace`, `preg_split`, and related PCRE builtins.

crate::php_cases! {
    preg_match_finds_substring => {
        r#"<?php
echo preg_match('/world/', 'hello world') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    preg_match_anchors_start_and_end => {
        r#"<?php
echo preg_match('/^abc$/', 'abc') ? 'full' : 'partial';
"#,
        ["full"]
    };

    preg_match_captures_group => {
        r#"<?php
preg_match('/(\d+)-(\d+)/', '12-34', $m);
echo $m[1] . ':' . $m[2];
"#,
        ["12:34"]
    };

    preg_match_all_counts_occurrences => {
        r#"<?php
echo preg_match_all('/a/', 'banana');
"#,
        ["3"]
    };

    preg_replace_literal_text => {
        r#"<?php
echo preg_replace('/cat/', 'dog', 'the cat sat');
"#,
        ["the dog sat"]
    };

    preg_replace_backreference => {
        r#"<?php
echo preg_replace('/(\w+) (\w+)/', '$2 $1', 'foo bar');
"#,
        ["bar foo"]
    };

    preg_replace_callback_transforms => {
        r#"<?php
echo preg_replace_callback('/\d+/', fn(array $m): string => (string)((int)$m[0] * 2), 'a3b');
"#,
        ["a6b"]
    };

    preg_split_on_delimiter => {
        r#"<?php
echo implode('|', preg_split('/\s+/', 'one two three'));
"#,
        ["one|two|three"]
    };

    preg_split_with_limit => {
        r#"<?php
echo count(preg_split('/,/', 'a,b,c,d', 2));
"#,
        ["2"]
    };

    preg_grep_filters_array => {
        r#"<?php
$out = preg_grep('/^[aeiou]/i', ['apple', 'dog', 'egg']);
echo implode(',', array_values($out));
"#,
        ["apple,egg"]
    };

    preg_quote_escapes_delimiters => {
        r#"<?php
$q = preg_quote('a.b?');
echo str_contains($q, '\\') ? 'quoted' : 'raw';
"#,
        ["quoted"]
    };

    preg_match_case_insensitive_flag => {
        r#"<?php
echo preg_match('/abc/i', 'AbC') ? 'ci' : 'cs';
"#,
        ["ci"]
    };

    preg_match_multiline_caret => {
        r#"<?php
echo preg_match('/^b/m', "a\nb\nc") ? 'm' : 's';
"#,
        ["m"]
    };

    preg_match_dotall_flag => {
        r#"<?php
echo preg_match('/a.b/s', "a\nb") ? 'dotall' : 'line';
"#,
        ["dotall"]
    };

    preg_replace_empty_pattern_returns_unchanged => {
        r#"<?php
echo preg_replace('//', 'x', 'hi');
"#,
        ["hi"]
    };

    preg_match_returns_zero_on_no_match => {
        r#"<?php
echo preg_match('/zzz/', 'abc');
"#,
        ["0"]
    };

    preg_match_offset_capture_includes_position => {
        r#"<?php
preg_match('/ab/', 'zzab', $m, PREG_OFFSET_CAPTURE);
echo $m[0][1];
"#,
        ["2"]
    };

    preg_split_no_empty_flag => {
        r#"<?php
echo count(preg_split('/,/', ',a,,b,', -1, PREG_SPLIT_NO_EMPTY));
"#,
        ["2"]
    };

    preg_replace_array_replacement => {
        r#"<?php
echo preg_replace(['/a/', '/b/'], ['1', '2'], 'ab');
"#,
        ["12"]
    };

    preg_match_named_capture_group => {
        r#"<?php
preg_match('/(?<year>\d{4})-(?<month>\d{2})/', '2024-06', $m);
echo $m['year'] . '-' . $m['month'];
"#,
        ["2024-06"]
    };

    preg_match_unicode_property => {
        r#"<?php
echo preg_match('/\p{L}+/u', 'café') ? 'letter' : 'no';
"#,
        ["letter"]
    };

    preg_split_delim_capture_flag => {
        r#"<?php
$parts = preg_split('/(:)/', 'a:b', -1, PREG_SPLIT_DELIM_CAPTURE);
echo implode('', $parts);
"#,
        ["a:b"]
    };

    preg_replace_limit_count => {
        r#"<?php
echo preg_replace('/a/', 'b', 'aaa', 2);
"#,
        ["bba"]
    };

    preg_match_all_set_order => {
        r#"<?php
preg_match_all('/(\d)(\d)/', '1234', $m, PREG_SET_ORDER);
echo count($m);
"#,
        ["2"]
    };

    preg_quote_custom_delimiter => {
        r#"<?php
echo preg_quote('a+b', '+');
"#,
        ["a\\+b"]
    };
}
