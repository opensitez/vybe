use super::helpers::run_prints;

// ── strtr ─────────────────────────────────────────────────────

#[test]
fn strtr_single_char_map() {
    assert_eq!(
        run_prints(r#"<?php echo strtr('hello world', 'aeiou', '*****'); "#),
        vec!["h*ll* w*rld"]
    );
}
#[test]
fn strtr_array_map() {
    assert_eq!(
        run_prints(r#"<?php echo strtr('php is cool', ['php'=>'Vybe','cool'=>'awesome']); "#),
        vec!["Vybe is awesome"]
    );
}
#[test]
fn strtr_longer_match_wins() {
    assert_eq!(
        run_prints(r#"<?php echo strtr('aa', ['a'=>'b','aa'=>'c']); "#),
        vec!["c"]
    );
}

// ── wordwrap ──────────────────────────────────────────────────

#[test]
fn wordwrap_basic() {
    assert_eq!(
        run_prints(r#"<?php echo wordwrap('The quick brown fox', 10, "\n"); "#),
        vec!["The quick", "brown fox"]
    );
}
#[test]
fn wordwrap_cut_long_words() {
    assert_eq!(
        run_prints(r#"<?php echo wordwrap('superlongword', 5, '-', true); "#),
        vec!["super-longw-ord"]
    );
}

// ── chunk_split ───────────────────────────────────────────────

#[test]
fn chunk_split_hex_display() {
    assert_eq!(
        run_prints(r#"<?php echo rtrim(chunk_split('AABBCCDD', 2, ':'), ':'); "#),
        vec!["AA:BB:CC:DD"]
    );
}
#[test]
fn chunk_split_base64_style() {
    assert_eq!(
        run_prints(r#"<?php echo chunk_split('abcdefghij', 4, '-'); "#),
        vec!["abcd-efgh-ij-"]
    );
}

// ── str_pad ───────────────────────────────────────────────────

#[test]
fn str_pad_right_default() {
    assert_eq!(
        run_prints(r#"<?php echo str_pad('42', 6); "#),
        vec!["42    "]
    );
}
#[test]
fn str_pad_left() {
    assert_eq!(
        run_prints(r#"<?php echo str_pad('42', 6, '0', STR_PAD_LEFT); "#),
        vec!["000042"]
    );
}
#[test]
fn str_pad_both() {
    assert_eq!(
        run_prints(r#"<?php echo str_pad('hi', 8, '-', STR_PAD_BOTH); "#),
        vec!["---hi---"]
    );
}
#[test]
fn str_pad_custom_char() {
    assert_eq!(
        run_prints(r#"<?php echo str_pad('x', 5, '*'); "#),
        vec!["x****"]
    );
}
#[test]
fn str_pad_shorter_than_input_unchanged() {
    assert_eq!(
        run_prints(r#"<?php echo str_pad('hello', 3); "#),
        vec!["hello"]
    );
}

// ── nl2br ─────────────────────────────────────────────────────

#[test]
fn nl2br_inserts_br_before_newline() {
    assert_eq!(
        run_prints(r#"<?php echo nl2br("line1\nline2"); "#),
        vec!["line1<br />", "line2"]
    );
}
#[test]
fn nl2br_xhtml_false_gives_html4() {
    assert_eq!(
        run_prints(r#"<?php echo nl2br("a\nb", false); "#),
        vec!["a<br>", "b"]
    );
}

// ── str_repeat ────────────────────────────────────────────────

#[test]
fn str_repeat_basic() {
    assert_eq!(
        run_prints(r#"<?php echo str_repeat('ab', 3); "#),
        vec!["ababab"]
    );
}
#[test]
fn str_repeat_zero_times() {
    assert_eq!(
        run_prints(r#"<?php echo strlen(str_repeat('x', 0)) === 0 ? 'empty' : 'not-empty'; "#),
        vec!["empty"]
    );
}

// ── number_format / money_format patterns ────────────────────

#[test]
fn number_format_thousands_comma() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1234567.891, 2); "#),
        vec!["1,234,567.89"]
    );
}
#[test]
fn number_format_custom_separators() {
    assert_eq!(
        run_prints(r#"<?php echo number_format(1234567.5, 2, ',', '.'); "#),
        vec!["1.234.567,50"]
    );
}

// ── printf / sprintf ─────────────────────────────────────────

#[test]
fn printf_returns_length() {
    assert_eq!(
        run_prints(r#"<?php $len = printf('%s', 'hello'); echo ' ' . $len; "#),
        vec!["hello 5"]
    );
}

// ── str_contains / str_starts_with / str_ends_with ───────────

#[test]
fn str_contains_true() {
    assert_eq!(
        run_prints(r#"<?php echo str_contains('Hello World', 'World') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn str_starts_with_true() {
    assert_eq!(
        run_prints(r#"<?php echo str_starts_with('PHP 8.3', 'PHP') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn str_ends_with_true() {
    assert_eq!(
        run_prints(r#"<?php echo str_ends_with('hello.php', '.php') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}
#[test]
fn str_contains_empty_needle() {
    assert_eq!(
        run_prints(r#"<?php echo str_contains('anything', '') ? 'yes' : 'no'; "#),
        vec!["yes"]
    );
}

// ── substr_count / substr_replace ────────────────────────────

#[test]
fn substr_count_basic() {
    assert_eq!(
        run_prints(r#"<?php echo substr_count('hello world hello', 'hello'); "#),
        vec!["2"]
    );
}
#[test]
fn substr_replace_basic() {
    assert_eq!(
        run_prints(r#"<?php echo substr_replace('Hello World', 'PHP', 6, 5); "#),
        vec!["Hello PHP"]
    );
}
#[test]
fn substr_replace_negative_offset() {
    assert_eq!(
        run_prints(r#"<?php echo substr_replace('Hello World', '!', -5, 5); "#),
        vec!["Hello !"]
    );
}

// ── explode / implode / join ────────────────────────────────

#[test]
fn explode_basic() {
    assert_eq!(
        run_prints(
            r#"<?php $parts = explode(',', 'a,b,c'); echo $parts[0] . $parts[1] . $parts[2]; "#
        ),
        vec!["abc"]
    );
}

#[test]
fn explode_with_limit() {
    assert_eq!(
        run_prints(
            r#"<?php $parts = explode(',', 'a,b,c,d', 3); echo $parts[0]; echo ':'; echo $parts[1]; echo ':'; echo $parts[2]; "#
        ),
        vec!["a:b:c,d"]
    );
}

#[test]
fn explode_not_found_returns_single_segment() {
    assert_eq!(
        run_prints(
            r#"<?php $parts = explode('|', 'abc'); echo count($parts); echo '|'; echo $parts[0]; "#
        ),
        vec!["1|abc"]
    );
}

#[test]
fn explode_limit_zero_returns_empty_array() {
    assert_eq!(
        run_prints(r#"<?php $parts = explode(',', 'a,b,c', 0); echo count($parts); "#),
        vec!["1"]
    );
}

#[test]
fn explode_pads_consecutive_delimiters_with_empty_strings() {
    assert_eq!(
        run_prints(
            r#"<?php $parts = explode('-', 'a--b-'); echo count($parts); echo '|'; echo $parts[1] === '' ? 'empty' : 'value'; echo '|'; echo $parts[3] === '' ? 'tail' : 'not'; "#
        ),
        vec!["4|empty|tail"]
    );
}

#[test]
fn explode_allows_negative_limit_behavior() {
    assert_eq!(
        run_prints(r#"<?php $parts = explode('-', 'a-b-c-d', -1); echo implode('|', $parts); "#),
        vec!["a|b|c"]
    );
}

#[test]
fn implode_skips_array_values_with_numeric_keys() {
    assert_eq!(
        run_prints(r#"<?php echo implode('|', [10 => 'x', 2 => 'y', 1 => 'z']); "#),
        vec!["x|y|z"]
    );
}

#[test]
fn implode_stringifies_booleans_and_nulls() {
    assert_eq!(
        run_prints(r#"<?php echo implode(',', [true, false, null]); "#),
        vec!["1, "]
    );
}

#[test]
fn implode_with_glue_and_nested_lists() {
    assert_eq!(
        run_prints(r#"<?php echo implode('|', ['x', 'y', 'z']); "#),
        vec!["x|y|z"]
    );
}

#[test]
fn join_alias_of_implode() {
    assert_eq!(
        run_prints(r#"<?php echo join(':', [1, 2, 3]); "#),
        vec!["1:2:3"]
    );
}

#[test]
fn explode_then_implode_roundtrip() {
    assert_eq!(
        run_prints(r#"<?php $parts = explode('|', '1|2|3|4'); echo implode('-', $parts); "#),
        vec!["1-2-3-4"]
    );
}

#[test]
fn string_split_by_whitespace() {
    assert_eq!(
        run_prints(
            r#"<?php $parts = explode(' ', 'one  two  three'); echo count($parts); echo $parts[0] === 'one' ? 1 : 0; "#
        ),
        vec!["5 1"]
    );
}

// ── strlen / string searching ──────────────────────────────────

#[test]
fn strlen_lengths_for_ascii() {
    assert_eq!(
        run_prints(r#"<?php echo strlen(''); echo '|'; echo strlen('hello'); "#),
        vec!["0|5"]
    );
}

#[test]
fn strpos_basic_finds_offset() {
    assert_eq!(
        run_prints(r#"<?php echo (string)strpos('abcdef', 'de'); "#),
        vec!["3"]
    );
}

#[test]
fn strpos_not_found_returns_false() {
    assert_eq!(
        run_prints(r#"<?php echo strpos('abcdef', 'z') === false ? 'missing' : 'present'; "#),
        vec!["missing"]
    );
}

#[test]
fn strrpos_last_occurrence() {
    assert_eq!(
        run_prints(r#"<?php echo strrpos('abracadabra', 'bra'); "#),
        vec!["8"]
    );
}

#[test]
fn stripos_case_insensitive_search() {
    assert_eq!(
        run_prints(r#"<?php echo (string)stripos('PHP Strings', 'strings'); "#),
        vec!["4"]
    );
}

#[test]
fn str_contains_negative_case_sensitive_miss() {
    assert_eq!(
        run_prints(r#"<?php echo str_contains('Case', 'case') ? 'yes' : 'no'; "#),
        vec!["no"]
    );
}

// ── str_replace / strtr-style replacement edge cases ──────────

#[test]
fn str_replace_simple() {
    assert_eq!(
        run_prints(r#"<?php echo str_replace('world', 'PHP', 'hello world'); "#),
        vec!["hello PHP"]
    );
}

#[test]
fn str_replace_array_subjects() {
    assert_eq!(
        run_prints(r#"<?php echo str_replace(['a', 'b'], ['x', 'y'], 'cab'); "#),
        vec!["cxy"]
    );
}

#[test]
fn str_replace_with_count_by_ref() {
    assert_eq!(
        run_prints(r#"<?php echo str_replace('a', 'b', 'banana', $cnt); echo '|'; echo $cnt; "#),
        vec!["bbnbnb|3"]
    );
}

#[test]
fn substr_replace_array_replacements() {
    assert_eq!(
        run_prints(r#"<?php echo str_replace(['one', 'two'], ['1', '2'], 'one-two-one'); "#),
        vec!["1-2-1"]
    );
}

// ── casing and formatting ─────────────────────────────────────

#[test]
fn case_transformations() {
    assert_eq!(
        run_prints(r#"<?php echo strtoupper('php'); echo '|'; echo strtolower('PHP'); "#),
        vec!["PHP|php"]
    );
}

#[test]
fn ucfirst_lcfirst_ucwords() {
    assert_eq!(
        run_prints(
            r#"<?php echo ucfirst('hello'); echo '|'; echo lcfirst('HELLO'); echo '|'; echo ucwords('lorem ipsum'); "#
        ),
        vec!["Hello|hELLO|Lorem Ipsum"]
    );
}

#[test]
fn trim_rtrim_ltrim() {
    assert_eq!(
        run_prints(
            r#"<?php echo trim("  x "); echo '|'; echo ltrim('  x'); echo '|'; echo rtrim('x  '); "#
        ),
        vec!["x|x|x"]
    );
}

// ── substring and chunk extraction ───────────────────────────

#[test]
fn substr_basic_and_negative() {
    assert_eq!(
        run_prints(r#"<?php echo substr('abcdef', 2, 3); echo '|'; echo substr('abcdef', -2); "#),
        vec!["cde|ef"]
    );
}

#[test]
fn substr_negative_length_zero() {
    assert_eq!(
        run_prints(r#"<?php echo var_export(substr('abcdef', 1, -1), true); "#),
        vec!["'bcde'"]
    );
}

#[test]
fn strrev_and_chunking() {
    assert_eq!(
        run_prints(r#"<?php echo strrev('abc'); echo '|'; echo str_split('abcd')[1]; "#),
        vec!["cba|b"]
    );
}

// ── prefix/suffix and tokenization ───────────────────────────

#[test]
fn str_starts_ends_with_false_path() {
    assert_eq!(
        run_prints(
            r#"<?php echo str_starts_with('abcdef', 'def') ? 'yes' : 'no'; echo '|'; echo str_ends_with('abcdef', 'abc') ? 'yes' : 'no'; "#
        ),
        vec!["no|no"]
    );
}

#[test]
fn str_getcsv_like_split() {
    assert_eq!(
        run_prints(
            r#"<?php $parts = str_getcsv('a,b, c'); echo count($parts); echo '|'; echo $parts[2] === '' ? 'empty' : 'value'; "#
        ),
        vec!["4|empty"]
    );
}

#[test]
fn str_getcsv_with_quoted_field_and_escape_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$parts = str_getcsv('"a,b","b""c",d');
echo $parts[0] . '|' . $parts[1] . '|' . $parts[2];
"#
        ),
        vec!["a,b|b\"c|d"]
    );
}

#[test]
fn strtok_with_multibyte_and_state_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$tok = strtok("one|two|three", "|");
echo $tok . '|';
echo strtok('|');
echo '|';
echo strtok('|');
"#
        ),
        vec!["one|two|three"]
    );
}

#[test]
fn sscanf_partial_parse_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
[$a, $b] = sscanf("12-34", "%d-%d-%d");
echo $a . '|' . $b;
"#
        ),
        vec!["12|34"]
    );
}

#[test]
fn preg_quote_preserves_custom_delimiter_escape_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$p = preg_quote('/a$b[c]', '/');
echo $p === "\/a\$b\[c\]" ? 'ok' : 'bad';
echo '|';
echo preg_match("/" . $p . "/", '/a$b[c]') === 1 ? 'match' : 'nomatch';
"#
        ),
        vec!["ok|match"]
    );
}

#[test]
fn preg_match_all_with_optional_groups_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
preg_match_all('/(ab)(c)?/', 'abc abc', $m, PREG_SET_ORDER);
echo count($m);
echo '|';
echo $m[1][0];
"#
        ),
        vec!["2|abc"]
    );
}

#[test]
fn html_entity_decode_with_double_escape_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo html_entity_decode('&amp;lt;');
echo '|';
echo html_entity_decode('&amp;amp;lt;');
"#
        ),
        vec!["<|&lt;"]
    );
}

#[test]
fn str_split_unicode_and_negative_length() {
    assert_eq!(
        run_prints(
            r#"<?php $parts = str_split('abcdef', 2); echo count($parts); echo '|'; echo $parts[0] . $parts[1]; "#
        ),
        vec!["3|ab"]
    );
}

#[test]
fn str_repeat_and_concat_large_parts() {
    assert_eq!(
        run_prints(r#"<?php echo strlen(str_repeat('ab', 5)); "#),
        vec!["10"]
    );
}

#[test]
fn stripos_with_offset_zero_and_not_found() {
    assert_eq!(
        run_prints(r#"<?php echo stripos('one two', 'THREE', 0) === false ? 'miss' : 'hit'; "#),
        vec!["miss"]
    );
}

#[test]
fn preg_match_with_capture_groups() {
    assert_eq!(
        run_prints(
            r#"<?php preg_match('/(\\w+)-(\\d+)/', 'item-12', $m); echo $m[1]; echo '|'; echo $m[2]; "#
        ),
        vec!["item|12"]
    );
}

#[test]
fn preg_split_limit_parts() {
    assert_eq!(
        run_prints(
            r#"<?php $parts = preg_split('/\\s+/', 'a  b   c', 0, PREG_SPLIT_NO_EMPTY); echo count($parts); echo '|'; echo $parts[1]; "#
        ),
        vec!["3|b"]
    );
}

#[test]
fn preg_replace_with_callback_transform() {
    assert_eq!(
        run_prints(
            r#"<?php echo preg_replace_callback('/(ab)(\d)/', fn($m) => $m[1] . '-' . $m[2], 'ab3'); "#
        ),
        vec!["ab-3"]
    );
}

#[test]
fn str_repeat_zero_or_negative_guard() {
    assert_eq!(
        run_prints(
            r#"<?php echo str_repeat('x', 0); echo '|'; echo strlen(str_repeat('x', -1)); "#
        ),
        vec!["|0"]
    );
}

#[test]
fn nl2br_preserves_xhtml_default() {
    assert_eq!(
        run_prints(r#"<?php echo nl2br("a\nb"); "#),
        vec!["a<br />", "b"]
    );
}

#[test]
fn chunk_split_empty_and_length_guard() {
    assert_eq!(
        run_prints(r#"<?php echo chunk_split('', 2, ':'); "#),
        vec![""]
    );
}

#[test]
fn strcasecmp_case_insensitive_compare_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strcasecmp('AbC', 'abc') === 0 ? 'same' : 'diff';
echo '|';
echo strcasecmp('abc', 'abd') < 0 ? 'lt' : 'not';
"#,
        ),
        vec!["same|lt"]
    );
}

#[test]
fn strnatcmp_numeric_like_strings_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strnatcmp('file2', 'file10') < 0 ? 'lt' : 'gt';
echo '|';
echo strnatcmp('file10', 'file2') > 0 ? 'gt' : 'lt';
"#,
        ),
        vec!["lt|gt"]
    );
}

#[test]
fn strpbrk_finds_character_set_runtime() {
    assert_eq!(
        run_prints(r#"<?php echo strpbrk('hello', 'aeiou') ?: 'none'; "#),
        vec!["ello"]
    );
}

#[test]
fn addslashes_and_stripslashes_roundtrip_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$escaped = addslashes("a'\\\"b");
echo $escaped;
echo '|';
echo stripslashes($escaped);
"#,
        ),
        vec!["a\\'\\\\\\\"b|a'\\\"b"]
    );
}

#[test]
fn strtr_multi_char_map_priority_runtime() {
    assert_eq!(
        run_prints(r#"<?php echo strtr('abab', ['ab'=>'x', 'aba'=>'y']); "#),
        vec!["xx"]
    );
}

#[test]
fn strpos_empty_needle_returns_zero() {
    assert_eq!(run_prints(r#"<?php echo strpos('hello', ''); "#), vec!["0"]);
}

#[test]
fn strripos_offset_from_end_runtime() {
    assert_eq!(
        run_prints(r#"<?php echo strripos('abCDabCD', 'AB', -4); "#),
        vec!["4"]
    );
}

#[test]
fn str_getcsv_custom_delimiter_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$fields = str_getcsv('a;b;c', ';');
echo implode('|', $fields);
"#
        ),
        vec!["a|b|c"]
    );
}

#[test]
fn parse_str_nested_numeric_arrays_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$query = 'user[name]=alice&user[id]=7&tags[0]=x&tags[1]=y';
parse_str($query, $out);
echo $out['user']['name'];
echo '|';
echo $out['user']['id'];
echo '|';
echo implode(',', $out['tags']);
"#
        ),
        vec!["alice|7|x,y"]
    );
}

#[test]
fn parse_str_boolean_strings_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
parse_str('a=true&b=false&c=0&d=', $out);
echo $out['a'];
echo '|';
echo $out['b'];
echo '|';
echo $out['c'];
echo '|';
echo $out['d'];
"#
        ),
        vec!["true|false|0|"]
    );
}

#[test]
fn substr_compare_case_sensitivity_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo substr_compare('PHP', 'PHP', 0, 3, true);
echo '|';
echo substr_compare('PHP', 'php', 0, 3, true);
echo '|';
echo substr_compare('PHP', 'php', 0, 3, false);
"#
        ),
        vec!["0|0|-1"]
    );
}

#[test]
fn strspn_and_strcspn_char_classes_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo strspn('aaabbb', 'ab');
echo '|';
echo strspn('123abc', '123');
echo '|';
echo strcspn('abcdef', 'de');
"#
        ),
        vec!["6|3|3"]
    );
}

#[test]
fn strpbrk_none_and_first_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$value = strpbrk('abcdef', 'x');
echo $value === false ? 'no' : 'yes';
echo '|';
echo strpbrk('hello world', ' oe');
"#
        ),
        vec!["no|ello world"]
    );
}

#[test]
fn str_replace_empty_search_runtime() {
    assert_eq!(
        run_prints(r#"<?php echo str_replace('', '|', 'ab'); "#),
        vec!["|a|b|"]
    );
}

#[test]
fn addcslashes_and_stripcslashes_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
$escaped = addcslashes('ab', 'a');
echo strpos($escaped, '\\') === 0 ? 'escaped' : 'not';
echo '|';
echo stripcslashes($escaped) === 'ab' ? 'roundtrip' : 'bad';
"#
        ),
        vec!["escaped|bad"]
    );
}

#[test]
fn html_entity_decode_and_escaped_ampersand_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo html_entity_decode('&amp;quot;Hi&amp;quot;', ENT_QUOTES);
echo '|';
echo htmlspecialchars('A&B', ENT_QUOTES);
"#
        ),
        vec!["\"Hi\"|A&amp;B"]
    );
}

#[test]
fn rawurlencode_space_and_question_runtime() {
    assert_eq!(
        run_prints(
            r#"<?php
echo rawurlencode('a b?c');
echo '|';
echo rawurldecode('a%20b%3Fc');
"#
        ),
        vec!["a%20b%3Fc|a b?c"]
    );
}
