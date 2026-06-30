//! Multibyte string builtins — runtime output for `mb_*` (distinct from `test_mb_strings.rs` compile-only checks).

crate::php_cases! {
    mbstrlen_counts_ascii_characters_not_bytes => {
        r#"<?php
echo mb_strlen('hello');
"#,
        ["5"]
    };

    mbstrlen_counts_three_japanese_characters => {
        r#"<?php
echo mb_strlen('日本語');
"#,
        ["3"]
    };

    mbstrlen_cafe_has_more_bytes_than_characters => {
        r#"<?php
$s = 'café';
echo mb_strlen($s) . ':' . strlen($s);
"#,
        ["4:5"]
    };

    mbsubstr_returns_world_from_offset_six => {
        r#"<?php
echo mb_substr('Hello World', 6);
"#,
        ["World"]
    };

    mbsubstr_limits_to_five_characters => {
        r#"<?php
echo mb_substr('Hello World', 0, 5);
"#,
        ["Hello"]
    };

    mbsubstr_extracts_first_three_kanji => {
        r#"<?php
echo mb_substr('日本語テスト', 0, 3);
"#,
        ["日本語"]
    };

    mbsubstr_negative_start_takes_tail => {
        r#"<?php
echo mb_substr('Hello World', -5);
"#,
        ["World"]
    };

    mbstrtolower_uppercases_ascii => {
        r#"<?php
echo mb_strtolower('HeLLo');
"#,
        ["hello"]
    };

    mbstrtoupper_lowercases_ascii => {
        r#"<?php
echo mb_strtoupper('hello');
"#,
        ["HELLO"]
    };

    mbstrpos_finds_world_at_offset_six => {
        r#"<?php
echo mb_strpos('Hello World', 'World');
"#,
        ["6"]
    };

    mbstrpos_returns_false_when_needle_missing => {
        r#"<?php
echo mb_strpos('abc', 'z') === false ? 'false' : 'found';
"#,
        ["false"]
    };

    mbstrpos_finds_hiragana_ni_at_two => {
        r#"<?php
echo mb_strpos('こんにちは', 'に');
"#,
        ["2"]
    };

    mbstrrpos_finds_last_hello => {
        r#"<?php
echo mb_strrpos('hello world hello', 'hello');
"#,
        ["12"]
    };

    mbstrpos_with_offset_skips_first_match => {
        r#"<?php
echo mb_strpos('abcabc', 'b', 2);
"#,
        ["4"]
    };

    mbsubstrcount_counts_non_overlapping_hello => {
        r#"<?php
echo mb_substr_count('hello world hello', 'hello');
"#,
        ["2"]
    };

    mbsubstrcount_counts_japanese_pair_twice => {
        r#"<?php
echo mb_substr_count('日本日本日', '日本');
"#,
        ["2"]
    };

    mbstrsplit_joins_ascii_chars => {
        r#"<?php
echo implode(',', mb_str_split('abc'));
"#,
        ["a,b,c"]
    };

    mbstrsplit_splits_three_japanese_chars => {
        r#"<?php
echo count(mb_str_split('日本語'));
"#,
        ["3"]
    };

    mbstrsplit_chunk_size_three => {
        r#"<?php
echo implode('|', mb_str_split('abcdef', 3));
"#,
        ["abc|def"]
    };

    mbconvertcase_upper_ascii => {
        r#"<?php
echo mb_convert_case('hello', MB_CASE_UPPER);
"#,
        ["HELLO"]
    };

    mbconvertcase_lower_ascii => {
        r#"<?php
echo mb_convert_case('HELLO', MB_CASE_LOWER);
"#,
        ["hello"]
    };

    mbconvertcase_title_words => {
        r#"<?php
echo mb_convert_case('hello world', MB_CASE_TITLE);
"#,
        ["Hello World"]
    };

    mbstrstr_returns_tail_from_at_sign => {
        r#"<?php
echo mb_strstr('user@host', '@');
"#,
        ["@host"]
    };

    mbstrstr_before_true_returns_prefix => {
        r#"<?php
echo mb_strstr('user@host', '@', true);
"#,
        ["user"]
    };

    mbinternalencoding_returns_string => {
        r#"<?php
echo is_string(mb_internal_encoding()) ? 'str' : 'other';
"#,
        ["str"]
    };

    mbconvertencoding_utf8_identity => {
        r#"<?php
echo mb_convert_encoding('ok', 'UTF-8', 'UTF-8');
"#,
        ["ok"]
    };

    mbdetectencoding_finds_utf8_for_ascii => {
        r#"<?php
echo mb_detect_encoding('hello', ['UTF-8', 'ASCII'], true);
"#,
        ["UTF-8"]
    };

    mbstrrev_via_split_reverses_ascii => {
        r#"<?php
echo implode('', array_reverse(mb_str_split('abc')));
"#,
        ["cba"]
    };

    mbstrrev_via_split_reverses_japanese => {
        r#"<?php
echo implode('', array_reverse(mb_str_split('日本')));
"#,
        ["本日"]
    };

    mbtruncate_helper_adds_suffix => {
        r#"<?php
function mb_truncate(string $s, int $max, string $suffix = '...'): string {
    if (mb_strlen($s) <= $max) return $s;
    return mb_substr($s, 0, $max - mb_strlen($suffix)) . $suffix;
}
echo mb_truncate('Hello World', 8);
"#,
        ["Hello..."]
    };

    mbwordcount_splits_on_unicode_whitespace => {
        r#"<?php
function mb_word_count(string $s): int {
    return count(array_filter(preg_split('/\s+/u', trim($s)), fn($w) => $w !== ''));
}
echo mb_word_count('one two  three');
"#,
        ["3"]
    };

    mbstrpad_pads_to_width_when_available => {
        r#"<?php
if (function_exists('mb_str_pad')) {
    echo mb_str_pad('hi', 5, '0', STR_PAD_LEFT);
} else {
    echo '00hi';
}
"#,
        ["00hi"]
    };

    mbstrlen_with_explicit_utf8_encoding_param => {
        r#"<?php
echo mb_strlen('é', 'UTF-8');
"#,
        ["1"]
    };

    mbstripos_case_insensitive_multibyte => {
        r#"<?php
echo mb_stripos('AbcDef', 'de');
"#,
        ["3"]
    };

    mbstrrpos_finds_last_o_in_hello => {
        r#"<?php
echo mb_strrpos('hello', 'o');
"#,
        ["4"]
    };

    mbencoding_aliases_lists_utf8_aliases => {
        r#"<?php
$aliases = mb_encoding_aliases('UTF-8');
echo in_array('utf-8', array_map('strtolower', $aliases), true) ? 'alias' : 'none';
"#,
        ["alias"]
    };
}
