//! `strpos`, `strrpos`, `strstr`, `substr`, `str_contains`, `str_starts_with`, `str_ends_with`.

crate::php_cases! {
    str_contains_finds_substring => {
        r#"<?php
echo str_contains('hello world', 'world') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    str_starts_with_prefix => {
        r#"<?php
echo str_starts_with('/api/users', '/api') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    str_ends_with_suffix => {
        r#"<?php
echo str_ends_with('file.php', '.php') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    strpos_returns_first_offset => {
        r#"<?php
echo strpos('banana', 'na');
"#,
        ["2"]
    };

    strrpos_returns_last_offset => {
        r#"<?php
echo strrpos('banana', 'na');
"#,
        ["4"]
    };

    strstr_returns_tail_after_needle => {
        r#"<?php
echo strstr('hello@world', '@');
"#,
        ["@world"]
    };

    substr_extracts_middle => {
        r#"<?php
echo substr('abcdef', 2, 3);
"#,
        ["cde"]
    };

    substr_negative_start_counts_from_end => {
        r#"<?php
echo substr('abcdef', -2);
"#,
        ["ef"]
    };

    strlen_byte_length => {
        r#"<?php
echo strlen('abc');
"#,
        ["3"]
    };

    str_repeat_duplicates_string => {
        r#"<?php
echo str_repeat('ab', 3);
"#,
        ["ababab"]
    };

    str_pad_left_align => {
        r#"<?php
echo str_pad('7', 4, '0', STR_PAD_LEFT);
"#,
        ["0007"]
    };

    strrev_reverses_bytes => {
        r#"<?php
echo strrev('abc');
"#,
        ["cba"]
    };

    strtolower_uppercase_ascii => {
        r#"<?php
echo strtolower('AbC');
"#,
        ["abc"]
    };

    strtoupper_lowercase_ascii => {
        r#"<?php
echo strtoupper('AbC');
"#,
        ["ABC"]
    };

    ucfirst_capitalizes_first_char => {
        r#"<?php
echo ucfirst('hello');
"#,
        ["Hello"]
    };

    lcfirst_lowercases_first_char => {
        r#"<?php
echo lcfirst('Hello');
"#,
        ["hello"]
    };

    str_replace_single_occurrence => {
        r#"<?php
echo str_replace('cat', 'dog', 'the cat');
"#,
        ["the dog"]
    };

    str_ireplace_case_insensitive => {
        r#"<?php
echo str_ireplace('ABC', 'x', 'abCdef');
"#,
        ["xdef"]
    };

    strncmp_compares_prefix => {
        r#"<?php
echo strncmp('abcdef', 'abcxyz', 3);
"#,
        ["0"]
    };

    strcmp_lexicographic_order => {
        r#"<?php
echo strcmp('b', 'a');
"#,
        ["1"]
    };

    chop_trims_trailing_chars => {
        r#"<?php
echo chop("hello\n\n");
"#,
        ["hello"]
    };

    wordwrap_inserts_breaks => {
        r#"<?php
echo str_contains(wordwrap('one two three four', 7), "\n") ? 'wrap' : 'flat';
"#,
        ["wrap"]
    };

    str_contains_empty_needle => {
        r#"<?php
echo str_contains('abc', '') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    strstr_with_before_needle => {
        r#"<?php
echo strstr('hello world', 'o', true);
"#,
        ["hell"]
    };

    strpos_not_found_is_false => {
        r#"<?php
echo var_export(strpos('abc', 'z'), true);
"#,
        ["false"]
    };

    strrpos_negative_offset => {
        r#"<?php
echo strrpos('ababa', 'ba', -4);
"#,
        ["3"]
    };

    str_starts_with_empty_needle_is_true => {
        r#"<?php
echo str_starts_with('hello', '') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    str_ends_with_empty_needle_is_true => {
        r#"<?php
echo str_ends_with('hello', '') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    str_contains_empty_haystack_is_false => {
        r#"<?php
echo str_contains('', 'a') ? 'yes' : 'no';
"#,
        ["no"]
    };

    strripos_from_offset_case_insensitive => {
        r#"<?php
echo strripos('Hello HeLLo', 'hello', 0);
"#,
        ["6"]
    };

    strpos_with_negative_offset_beyond_start => {
        r#"<?php
echo strpos('abcdef', 'a', -10);
"#,
        ["0"]
    };

    strstr_returns_false_for_missing => {
        r#"<?php
echo var_export(strstr('abc', 'z'), true);
"#,
        ["false"]
    };

    strncmp_with_length_greater_than_lengths => {
        r#"<?php
echo strncmp('abc', 'abcd', 99) < 0 ? 'lt' : 'ge';
"#,
        ["lt"]
    };

    substr_zero_length => {
        r#"<?php
echo var_export(substr('abcdef', 2, 0), true);
"#,
        ["''"]
    };

    strcmp_zero_sign_and_case => {
        r#"<?php
echo strcmp('abc', 'abc') === 0 ? 'eq' : 'ne';
echo '|';
echo strcmp('abc', 'abd');
"#,
        ["eq|-1"]
    };
}
