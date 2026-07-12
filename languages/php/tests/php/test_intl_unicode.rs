//! `grapheme_*`, `Normalizer`, `Collator`, and Unicode-aware `preg` — intl/UTF-8 behaviors.

crate::php_cases! {
    grapheme_strlen_counts_emoji_as_one => {
        r#"<?php
if (!function_exists('grapheme_strlen')) { echo 'skip'; } else {
    echo grapheme_strlen('😀');
}
"#,
        ["1"]
    };

    grapheme_substr_extracts_first_cluster => {
        r#"<?php
if (!function_exists('grapheme_substr')) { echo 'skip'; } else {
    echo grapheme_substr('日本語', 0, 2);
}
"#,
        ["日本"]
    };

    grapheme_strpos_finds_second_cluster => {
        r#"<?php
if (!function_exists('grapheme_strpos')) { echo 'skip'; } else {
    echo grapheme_strpos('日本語', '語');
}
"#,
        ["2"]
    };

    normalizer_nfc_shortens_composed_sequence => {
        r#"<?php
if (!class_exists('Normalizer')) { echo 'skip'; } else {
    $d = Normalizer::normalize("e\u{0301}", Normalizer::FORM_C);
    echo strlen($d) < strlen("e\u{0301}") ? 'nfc' : 'same';
}
"#,
        ["nfc"]
    };

    normalizer_is_normalized_detects_nfc => {
        r#"<?php
if (!class_exists('Normalizer')) { echo 'skip'; } else {
    echo Normalizer::isNormalized('café', Normalizer::FORM_C) ? 'yes' : 'no';
}
"#,
        ["yes"]
    };

    collator_compare_orders_fruits_in_french_locale => {
        r#"<?php
if (!class_exists('Collator')) { echo 'skip'; } else {
    $c = new Collator('fr_FR');
    echo $c->compare('apple', 'banana') < 0 ? 'before' : 'after';
}
"#,
        ["before"]
    };

    collator_sort_sorts_utf8_array => {
        r#"<?php
if (!class_exists('Collator')) { echo 'skip'; } else {
    $a = ['č', 'a', 'b'];
    $c = new Collator('root');
    $c->sort($a);
    echo $a[0];
}
"#,
        ["a"]
    };

    intlcal_get_now_returns_positive_timestamp => {
        r#"<?php
if (!class_exists('IntlCalendar')) { echo 'skip'; } else {
    echo IntlCalendar::getNow() > 0 ? 'now' : 'zero';
}
"#,
        ["now"]
    };

    preg_match_unicode_letter_property => {
        r#"<?php
echo preg_match('/\p{L}/u', '字') ? 'letter' : 'no';
"#,
        ["letter"]
    };

    preg_match_unicode_digit_property => {
        r#"<?php
echo preg_match('/\p{N}/u', '７') ? 'digit' : 'no';
"#,
        ["digit"]
    };

    preg_replace_unicode_case_insensitive => {
        r#"<?php
echo preg_replace('/über/ui', 'X', 'ÜBER');
"#,
        ["X"]
    };

    json_encode_invalid_utf8_substitute_replaces_bytes => {
        r#"<?php
$out = json_encode("\xB1\x31", JSON_INVALID_UTF8_SUBSTITUTE);
echo str_contains($out, '1') ? 'substituted' : 'lost';
"#,
        ["substituted"]
    };

    htmlspecialchars_ent_substitute_replaces_invalid_utf8 => {
        r#"<?php
$out = htmlspecialchars("\xC3\x28", ENT_SUBSTITUTE | ENT_QUOTES, 'UTF-8');
echo strlen($out) > 0 ? 'out' : 'empty';
"#,
        ["out"]
    };
}
