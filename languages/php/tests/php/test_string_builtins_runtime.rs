//! Runtime output for formatting builtins not covered by `test_string_formatting.rs` `run_prints` cases.

crate::php_cases! {
    number_format_thousands_and_decimals => {
        r#"<?php
echo number_format(9876543.21, 2, '.', ',');
"#,
        ["9,876,543.21"]
    };

    number_format_zero_decimals_rounds => {
        r#"<?php
echo number_format(1000);
"#,
        ["1,000"]
    };

    vsprintf_builds_string_from_argument_array => {
        r#"<?php
echo vsprintf('%s:%d', ['item', 7]);
"#,
        ["item:7"]
    };

    sscanf_parses_iso_date_components => {
        r#"<?php
[$y, $m, $d] = sscanf('2024-07-15', '%d-%d-%d');
echo "$y-$m-$d";
"#,
        ["2024-7-15"]
    };

    strip_tags_removes_script_keeps_allowed => {
        r#"<?php
echo strip_tags('<p>ok</p><script>x</script>', '<p>');
"#,
        // PHP strips the <script> tags but keeps their text content ("x").
        ["<p>ok</p>x"]
    };

    htmlspecialchars_encodes_angle_and_amp => {
        r#"<?php
echo htmlspecialchars('<a&>');
"#,
        ["&lt;a&amp;&gt;"]
    };

    htmlspecialchars_decode_roundtrip => {
        r#"<?php
$s = '<b>x</b>';
echo htmlspecialchars_decode(htmlspecialchars($s));
"#,
        ["<b>x</b>"]
    };

    html_entity_decode_named_entities => {
        r#"<?php
echo html_entity_decode('&lt;b&gt;');
"#,
        ["<b>"]
    };

    addslashes_escapes_quotes => {
        r#"<?php
echo addslashes("a'b\"c");
"#,
        ["a\\'b\\\"c"]
    };

    stripslashes_unescapes_quotes => {
        r#"<?php
echo stripslashes("a\\'b");
"#,
        ["a'b"]
    };

    quotemeta_prefixes_regex_metachars => {
        r#"<?php
echo quotemeta('.?*');
"#,
        ["\\.\\?\\*"]
    };

    chunk_split_inserts_separator_every_three => {
        r#"<?php
echo chunk_split('abcdef', 3, '-');
"#,
        ["abc-def-"]
    };

    wordwrap_breaks_long_line => {
        r#"<?php
echo str_contains(wordwrap('one two three four', 7), "\n") ? 'wrapped' : 'flat';
"#,
        ["wrapped"]
    };

    str_word_count_counts_three_words => {
        r#"<?php
echo str_word_count('one two three');
"#,
        ["3"]
    };

    nl2br_inserts_br_before_newline => {
        r#"<?php
echo nl2br("a\nb", false);
"#,
        ["a<br>", "b"]
    };

    str_getcsv_parses_quoted_comma_field => {
        r#"<?php
$row = str_getcsv('"a,b",c');
echo $row[0] . '|' . $row[1];
"#,
        ["a,b|c"]
    };

    sprintf_zero_padded_width => {
        r#"<?php
echo sprintf('%04d', 7);
"#,
        ["0007"]
    };

    str_rot13_rotates_letters => {
        r#"<?php
echo str_rot13('abc');
"#,
        ["nop"]
    };

    convert_uuencode_then_decode_roundtrip => {
        r#"<?php
echo convert_uudecode(convert_uuencode('data'));
"#,
        ["data"]
    };

    localeconv_decimal_point_is_string => {
        r#"<?php
$lc = localeconv();
echo is_string($lc['decimal_point']) ? 'dp' : 'no';
"#,
        ["dp"]
    };

    number_format_negative_value => {
        r#"<?php
echo number_format(-1234.56, 2, '.', ',');
"#,
        ["-1,234.56"]
    };

    number_format_zero_pad_width => {
        r#"<?php
echo sprintf('%08d', 1234);
"#,
        ["00001234"]
    };

    vsprintf_supports_multiple_types => {
        r#"<?php
echo vsprintf('%s-%d-%.2f', ['item', 7, 2.5]);
"#,
        ["item-7-2.50"]
    };

    sscanf_with_mismatched_input_still_parses_prefix => {
        r#"<?php
[$a, $b] = sscanf('1:2:3', '%d:%d');
echo "$a|$b";
"#,
        ["1|2"]
    };

    htmlspecialchars_disables_double_encode => {
        r#"<?php
echo htmlspecialchars('<b>safe</b>', ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8', false);
"#,
        ["&lt;b&gt;safe&lt;/b&gt;"]
    };

    html_entity_decode_double_quotes => {
        r#"<?php
echo html_entity_decode('&quot;x&quot;');
"#,
        ["\"x\""]
    };

    quotemeta_escapes_slash_chars => {
        r#"<?php
echo quotemeta('a/b|c');
"#,
        ["a/b\\|c"]
    };

    nl2br_with_xhtml_true => {
        r#"<?php
echo nl2br("x\n", true);
"#,
        ["x<br />\n"]
    };

    number_format_custom_thousands_sep => {
        r#"<?php
echo number_format(1234567.89, 2, ',', '.');
"#,
        ["1.234.567,89"]
    };

    sscanf_parses_prefixed_sign => {
        r#"<?php
[$a, $b] = sscanf('+007:abc', '%d:%s');
echo $a . '|' . $b;
"#,
        ["7|abc"]
    };

    str_repeat_zero => {
        r#"<?php
echo str_repeat('x', 0) === '' ? 'empty' : 'filled';
"#,
        ["empty"]
    };

    stripcslashes_addslashes_roundtrip => {
        r#"<?php
$s = addslashes("a\\nb\\'c\\\"d");
echo stripcslashes($s);
"#,
        ["a\nb'cd"]
    };

    html_entity_decode_without_double_quotes => {
        r#"<?php
echo html_entity_decode('&quot;x&quot;', ENT_COMPAT, 'UTF-8');
"#,
        ["\"x\""]
    };

    number_format_uses_optional_precision => {
        r#"<?php
echo number_format(12.34567, 3, ',', '');
"#,
        ["12,346"]
    };

    str_word_count_mode2_returns_positions => {
        r#"<?php
$m = str_word_count('ab cd ef', 2);
echo implode(',', array_keys($m));
"#,
        ["0,3,6"]
    };

    quoted_printable_encoded_decode_roundtrip => {
        r#"<?php
$encoded = quoted_printable_encode('Cafe: Crème');
$decoded = quoted_printable_decode($encoded);
echo strpos($encoded, '=') !== false ? 'quoted' : 'plain';
echo '|';
echo $decoded;
"#,
        ["quoted|Cafe: Crème"]
    };

    urlencode_and_rawurlencode_different_spaces => {
        r#"<?php
echo urlencode('a b');
echo '|';
echo rawurlencode('a b');
"#,
        ["a+b|a%20b"]
    };
}
