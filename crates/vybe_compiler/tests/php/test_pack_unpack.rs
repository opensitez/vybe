//! `pack`, `unpack`, `chr`, and `ord` — binary layout and byte conversions.

crate::php_cases! {
    ord_returns_ascii_code_for_capital_a => {
        r#"<?php
echo ord('A');
"#,
        ["65"]
    };

    chr_returns_character_from_code => {
        r#"<?php
echo chr(97);
"#,
        ["a"]
    };

    pack_unsigned_char_produces_single_byte => {
        r#"<?php
echo pack('C', 65);
"#,
        ["A"]
    };

    pack_three_chars_forms_hel => {
        r#"<?php
echo pack('CCC', 72, 101, 108);
"#,
        ["Hel"]
    };

    pack_big_endian_long_hex => {
        r#"<?php
echo bin2hex(pack('N', 16909060));
"#,
        ["01020304"]
    };

    pack_little_endian_long_hex => {
        r#"<?php
echo bin2hex(pack('V', 16909060));
"#,
        ["04030201"]
    };

    unpack_big_endian_long_roundtrip => {
        r#"<?php
$p = pack('N', 305419896);
echo unpack('N', $p)[1];
"#,
        ["305419896"]
    };

    pack_unsigned_short_big_endian_hex => {
        r#"<?php
echo bin2hex(pack('n', 0x0102));
"#,
        ["0102"]
    };

    pack_float_is_four_bytes => {
        r#"<?php
echo strlen(pack('f', 3.14));
"#,
        ["4"]
    };

    pack_double_is_eight_bytes => {
        r#"<?php
echo strlen(pack('d', 3.14159));
"#,
        ["8"]
    };

    pack_null_padding_character => {
        r#"<?php
echo strlen(pack('x'));
"#,
        ["1"]
    };

    unpack_multiple_format_codes => {
        r#"<?php
$data = pack('C2', 1, 2);
$r = unpack('Cfirst/Csecond', $data);
echo $r['first'] . ':' . $r['second'];
"#,
        ["1:2"]
    };

    chr_ord_roundtrip_for_printable => {
        r#"<?php
echo chr(ord('Z'));
"#,
        ["Z"]
    };

    pack_hex_format_emits_binary => {
        r#"<?php
echo bin2hex(pack('H*', '4142'));
"#,
        ["4142"]
    };
}
