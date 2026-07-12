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
}
