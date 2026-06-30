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
        ["<p>ok</p>"]
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
        ["a<br />\nb"]
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
}
