//! String comparison and phonetic helpers — runtime output (distinct from `test_string_advanced.rs`).

crate::php_cases! {
    strcmp_equal_returns_zero => {
        r#"<?php
echo strcmp('abc', 'abc');
"#,
        ["0"]
    };

    strcmp_less_returns_negative => {
        r#"<?php
echo strcmp('a', 'b') < 0 ? 'lt' : 'ge';
"#,
        ["lt"]
    };

    strcasecmp_case_insensitive_equal => {
        r#"<?php
echo strcasecmp('AbC', 'aBc');
"#,
        ["0"]
    };

    strnatcmp_orders_numbers_naturally => {
        r#"<?php
echo strnatcmp('img2', 'img10') < 0 ? 'before' : 'after';
"#,
        ["before"]
    };

    strnatcasecmp_natural_case_insensitive => {
        r#"<?php
echo strnatcasecmp('Img2', 'img10') < 0 ? 'before' : 'after';
"#,
        ["before"]
    };

    levenshtein_with_custom_costs => {
        r#"<?php
echo levenshtein('ca', 'abc', 1, 10, 10);
"#,
        ["12"]
    };

    levenshtein_insert_only => {
        r#"<?php
echo levenshtein('', 'abc');
"#,
        ["3"]
    };

    similar_text_percent_by_reference => {
        r#"<?php
similar_text('abcdef', 'abcxyz', $pct);
echo (int)$pct;
"#,
        ["50"]
    };

    metaphone_english_word => {
        r#"<?php
echo metaphone('program');
"#,
        ["PRKRM"]
    };

    metaphone_with_length_limit => {
        r#"<?php
echo metaphone('programming', 4);
"#,
        ["PRKR"]
    };

    soundex_same_sound => {
        r#"<?php
echo soundex('Smith') === soundex('Smyth') ? 'match' : 'diff';
"#,
        ["match"]
    };

    soundex_known_code => {
        r#"<?php
echo soundex('Euler');
"#,
        ["E460"]
    };

    strncmp_length_limited => {
        r#"<?php
echo strncmp('hello', 'help', 3);
"#,
        ["0"]
    };

    strncasecmp_length_limited => {
        r#"<?php
echo strncasecmp('Hello', 'help', 3);
"#,
        ["0"]
    };

    substr_compare_with_offset => {
        r#"<?php
echo substr_compare('abcdef', 'cde', 2, 3);
"#,
        ["0"]
    };

    substr_compare_case_insensitive => {
        r#"<?php
echo substr_compare('abcdef', 'CDE', 2, 3, true);
"#,
        ["0"]
    };

    strcoll_locale_c_default => {
        r#"<?php
echo strcoll('a', 'b') < 0 ? 'lt' : 'ge';
"#,
        ["lt"]
    };

    count_chars_frequency_mode => {
        r#"<?php
$m = count_chars('aab', 1);
echo $m[ord('a')];
"#,
        ["2"]
    };

    count_chars_unique_letters => {
        r#"<?php
echo implode('', array_keys(count_chars('aba', 3)));
"#,
        ["ab"]
    };

    strspn_accept_set => {
        r#"<?php
echo strspn('123abc', '0123456789');
"#,
        ["3"]
    };

    strcspn_reject_set => {
        r#"<?php
echo strcspn('123abc', '0123456789');
"#,
        ["0"]
    };

    str_word_count_default => {
        r#"<?php
echo str_word_count('one two three');
"#,
        ["3"]
    };

    str_word_count_with_array_return => {
        r#"<?php
echo count(str_word_count('a b c', 2));
"#,
        ["3"]
    };

    localeconv_decimal_point => {
        r#"<?php
$l = localeconv();
echo isset($l['decimal_point']) ? 'ok' : 'no';
"#,
        ["ok"]
    };

    str_rot13_roundtrip => {
        r#"<?php
echo str_rot13(str_rot13('hello'));
"#,
        ["hello"]
    };

    quoted_printable_encode_decode => {
        r#"<?php
$s = "a\r\nb";
echo quoted_printable_decode(quoted_printable_encode($s));
"#,
        // Roundtrip yields "a\r\nb"; the test harness splits stdout on '\n'
        // (and trims '\r'), so a correct result is captured as two lines.
        ["a", "b"]
    };

    convert_uuencode_decode_roundtrip => {
        r#"<?php
$enc = convert_uuencode("hi");
echo convert_uudecode($enc);
"#,
        ["hi"]
    };

    crc32_checksum => {
        r#"<?php
echo strlen(dechex(crc32('test'))) > 0 ? 'hex' : 'no';
"#,
        ["hex"]
    };

    md5_file_from_memory_stream => {
        r#"<?php
$f = fopen('php://memory', 'r+');
fwrite($f, 'data');
rewind($f);
$path = stream_get_meta_data($f)['uri'];
echo strlen(md5_file($path));
"#,
        ["32"]
    };

    hash_equals_timing_safe => {
        r#"<?php
echo hash_equals('abc', 'abc') ? 'eq' : 'ne';
"#,
        ["eq"]
    };

    str_starts_with_unicode_bytes => {
        r#"<?php
echo str_starts_with('café', 'caf') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    str_ends_with_suffix => {
        r#"<?php
echo str_ends_with('filename.txt', '.txt') ? 'yes' : 'no';
"#,
        ["yes"]
    };

    str_contains_substring => {
        r#"<?php
echo str_contains('haystack', 'needle') ? 'yes' : 'no';
"#,
        ["no"]
    };

    str_increment_alphanumeric => {
        r#"<?php
echo str_increment('a9');
"#,
        ["b0"]
    };

    str_decrement_alphanumeric => {
        r#"<?php
echo str_decrement('b0');
"#,
        ["a9"]
    };
}
