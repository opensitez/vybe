//! Additional `mb_*` APIs not covered by `test_mbstring.rs`, `test_mb_strings.rs`, or `test_string_case_multibyte.rs`.

crate::php_cases! {
    mbchr_returns_unicode_character_from_codepoint => {
        r#"<?php
echo mb_chr(0x4E16);
"#,
        ["世"]
    };

    mbord_returns_codepoint_for_unicode_char => {
        r#"<?php
echo mb_ord('世');
"#,
        ["19990"]
    };

    mbchr_mbord_roundtrip_for_eacute => {
        r#"<?php
echo mb_chr(mb_ord('é'));
"#,
        ["é"]
    };

    mbconvertkana_hiragana_to_katakana => {
        r#"<?php
echo mb_convert_kana('あいう', 'K');
"#,
        ["アイウ"]
    };

    mbconvertkana_fullwidth_latin_to_halfwidth => {
        r#"<?php
echo mb_convert_kana('ＡＢＣ', 'a');
"#,
        ["ABC"]
    };

    mbstrimwidth_truncates_by_display_width => {
        r#"<?php
echo mb_strimwidth('abcdef', 0, 4, '...');
"#,
        ["abc..."]
    };

    mbstrimwidth_respects_multibyte_width => {
        r#"<?php
echo mb_strimwidth('日本語', 0, 4, '..');
"#,
        ["日本.."]
    };

    mbstrwidth_counts_wide_characters_wider_than_strlen => {
        r#"<?php
$s = '日本';
echo mb_strwidth($s) . ':' . strlen($s);
"#,
        ["4:6"]
    };

    mbsplit_splits_on_utf8_delimiter => {
        r#"<?php
echo implode('|', mb_split('・', 'a・b・c'));
"#,
        ["a|b|c"]
    };

    mbstristr_finds_case_insensitive_multibyte_needle => {
        r#"<?php
echo mb_stristr('AbCdef', 'BC');
"#,
        ["Bcdef"]
    };

    mbstripos_finds_case_insensitive_offset => {
        r#"<?php
echo mb_stripos('Hello', 'LL');
"#,
        ["2"]
    };

    mblistencodings_includes_utf8 => {
        r#"<?php
echo in_array('UTF-8', mb_list_encodings(), true) ? 'utf8' : 'missing';
"#,
        ["utf8"]
    };

    mbpreferredmimename_utf8 => {
        r#"<?php
echo mb_preferred_mime_name('UTF-8');
"#,
        ["UTF-8"]
    };

    mblanguage_set_and_read_back => {
        r#"<?php
$old = mb_language('uni');
echo mb_language();
mb_language($old);
"#,
        ["uni"]
    };

    mbregexencoding_set_utf8 => {
        r#"<?php
$old = mb_regex_encoding('UTF-8');
echo mb_regex_encoding();
mb_regex_encoding($old);
"#,
        ["UTF-8"]
    };

    mbencodenumericentity_then_decode_roundtrip => {
        r#"<?php
$map = [0x80, 0xff, 0, 0xff];
$enc = mb_encode_numericentity('x', $map, 'UTF-8');
echo mb_decode_numericentity($enc, $map, 'UTF-8');
"#,
        ["x"]
    };

    mbsubstitutecharacter_none_disables_substitution => {
        r#"<?php
$old = mb_substitute_character();
mb_substitute_character('none');
echo mb_substitute_character() === 'none' ? 'none' : 'other';
mb_substitute_character($old);
"#,
        ["none"]
    };

    mbcheckencoding_rejects_invalid_utf8_sequence => {
        r#"<?php
echo mb_check_encoding("\xFF\xFE", 'UTF-8') ? 'valid' : 'invalid';
"#,
        ["invalid"]
    };

    mbconvertencoding_latin1_bytes_to_utf8 => {
        r#"<?php
$latin = "\xE9";
echo mb_convert_encoding($latin, 'UTF-8', 'ISO-8859-1');
"#,
        ["é"]
    };

    mbstrlen_emoji_counts_grapheme_clusters_php_style => {
        r#"<?php
echo mb_strlen('a😀b');
"#,
        ["3"]
    };

    mbstrsplit_preserves_emoji_as_single_cell => {
        r#"<?php
echo count(mb_str_split('a😀'));
"#,
        ["2"]
    };

    mb_get_info_returns_map => {
        r#"<?php
$info = mb_get_info();
echo is_array($info) ? 'array' : 'none';
echo '|';
echo $info['internal_encoding'] === mb_internal_encoding() ? 'ok' : 'bad';
"#,
        ["array|ok"]
    };

    mb_detect_order_set_and_restore => {
        r#"<?php
$old = mb_detect_order();
mb_detect_order(['ASCII', 'UTF-8']);
$new = mb_detect_order();
echo in_array('ASCII', $new, true) ? 'set' : 'unset';
mb_detect_order($old);
echo '|';
$restored = mb_detect_order();
echo is_array($restored) ? 'restored' : 'bad';
"#,
        ["set|restored"]
    };

    mb_scrub_invalid_bytes => {
        r#"<?php
$scrubbed = mb_scrub("\xFF", 'UTF-8');
echo $scrubbed === "\xEF\xBF\xBD" ? 'replaced' : 'other';
"#,
        ["replaced"]
    };
}
