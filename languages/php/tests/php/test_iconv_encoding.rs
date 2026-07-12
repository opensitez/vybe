//! `iconv` and cross-encoding conversion — distinct from the single `iconv` case in `test_types.rs`.

crate::php_cases! {
    iconv_utf8_to_utf8_identity => {
        r#"<?php
echo iconv('UTF-8', 'UTF-8', 'café');
"#,
        ["café"]
    };

    iconv_iso8859_1_bytes_to_utf8 => {
        r#"<?php
echo iconv('ISO-8859-1', 'UTF-8', "\xE9");
"#,
        ["é"]
    };

    iconv_utf8_to_iso8859_1_preserves_latin1_char => {
        r#"<?php
$out = iconv('UTF-8', 'ISO-8859-1', 'é');
echo bin2hex($out);
"#,
        ["e9"]
    };

    iconv_ignore_drops_invalid_sequences => {
        r#"<?php
echo iconv('UTF-8', 'UTF-8//IGNORE', "ok\xFF\xFEbad");
"#,
        ["okbad"]
    };

    iconv_translit_approximates_unrepresentable => {
        r#"<?php
$out = iconv('UTF-8', 'ASCII//TRANSLIT', 'über');
echo strlen($out) > 0 ? 'translit' : 'empty';
"#,
        ["translit"]
    };

    iconv_mime_decode_encoded_word => {
        r#"<?php
if (!function_exists('iconv_mime_decode')) { echo 'skip'; } else {
    echo iconv_mime_decode('=?UTF-8?B?Y2Fm6Q==?=', 0, 'UTF-8');
}
"#,
        ["café"]
    };

    iconv_mime_encode_produces_encoded_word_prefix => {
        r#"<?php
if (!function_exists('iconv_mime_encode')) { echo 'skip'; } else {
    $h = iconv_mime_encode('Subject', 'café', ['input-charset' => 'UTF-8', 'output-charset' => 'UTF-8']);
    echo str_starts_with($h, 'Subject: =?UTF-8?') ? 'mime' : 'raw';
}
"#,
        ["mime"]
    };

    mbconvertencoding_auto_detects_utf8_for_valid_sequence => {
        r#"<?php
echo mb_convert_encoding('日本', 'UTF-8', 'auto');
"#,
        ["日本"]
    };

    utf8_encode_decode_latin1_roundtrip_when_available => {
        r#"<?php
if (!function_exists('utf8_encode')) { echo 'skip'; } else {
    echo utf8_decode(utf8_encode("\xE9"));
}
"#,
        ["é"]
    };

    hex_binary_roundtrip_for_utf8_bytes => {
        r#"<?php
$bytes = hex2bin('c3a9');
echo mb_convert_encoding($bytes, 'UTF-8', 'UTF-8');
"#,
        ["é"]
    };
}
