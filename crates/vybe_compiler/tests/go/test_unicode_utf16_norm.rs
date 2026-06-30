//! unicode/utf16: Encode/Decode and surrogate pairs; unicode: IsLetter, IsDigit,
//! ToUpper/ToLower, SimpleFold; utf8: Valid, ValidRune, RuneLen, EncodeRune/DecodeRune
//! — extended coverage distinct from `test_unicode_package.rs` and `test_unicode_utf8.rs`.

use crate::helpers::*;

go_run_cases! {
    utf16_encode_ascii_runes => (
        "package main; import \"fmt\"; import \"unicode/utf16\"; func main() { u := utf16.Encode([]rune(\"AB\")); fmt.Println(len(u)); fmt.Println(int(u[0])); fmt.Println(int(u[1])) }",
        vec!["2", "65", "66"]
    ),
    utf16_encode_bmp_rune => (
        "package main; import \"fmt\"; import \"unicode/utf16\"; func main() { u := utf16.Encode([]rune(\"世\")); fmt.Println(len(u)); fmt.Println(int(u[0])) }",
        vec!["1", "19990"]
    ),
    utf16_encode_emoji_surrogate_pair => (
        "package main; import \"fmt\"; import \"unicode/utf16\"; func main() { u := utf16.Encode([]rune(\"🙂\")); fmt.Println(len(u)); fmt.Println(int(u[0])); fmt.Println(int(u[1])) }",
        vec!["2", "55357", "56898"]
    ),
    utf16_decode_ascii_units => (
        "package main; import \"fmt\"; import \"unicode/utf16\"; func main() { rs := utf16.Decode([]uint16{65, 66}); fmt.Println(string(rs)) }",
        vec!["AB"]
    ),
    utf16_decode_surrogate_pair_emoji => (
        "package main; import \"fmt\"; import \"unicode/utf16\"; func main() { rs := utf16.Decode([]uint16{0xD83D, 0xDE42}); fmt.Println(len(rs)); fmt.Println(int(rs[0])) }",
        vec!["1", "128578"]
    ),
    utf16_decode_replaces_invalid_surrogate => (
        "package main; import \"fmt\"; import \"unicode/utf16\"; func main() { rs := utf16.Decode([]uint16{0xD800}); fmt.Println(int(rs[0])) }",
        vec!["65533"]
    ),
    utf16_encode_rune_ascii => (
        "package main; import \"fmt\"; import \"unicode/utf16\"; func main() { u1, u2 := utf16.EncodeRune('Z'); fmt.Println(u1); fmt.Println(u2) }",
        vec!["90", "65535"]
    ),
    utf16_encode_rune_supplementary => (
        "package main; import \"fmt\"; import \"unicode/utf16\"; func main() { u1, u2 := utf16.EncodeRune(0x1F600); fmt.Println(int(u1)); fmt.Println(int(u2)) }",
        vec!["55296", "56832"]
    ),
    utf16_decode_rune_single_unit => (
        "package main; import \"fmt\"; import \"unicode/utf16\"; func main() { r := utf16.DecodeRune(65, 65535); fmt.Println(int(r)) }",
        vec!["65"]
    ),
    utf16_decode_rune_pair => (
        "package main; import \"fmt\"; import \"unicode/utf16\"; func main() { r := utf16.DecodeRune(0xD83D, 0xDE00); fmt.Println(int(r)) }",
        vec!["128512"]
    ),
    utf16_is_surrogate_high => (
        "package main; import \"fmt\"; import \"unicode/utf16\"; func main() { fmt.Println(utf16.IsSurrogate(0xD800)) }",
        vec!["true"]
    ),
    utf16_is_surrogate_low => (
        "package main; import \"fmt\"; import \"unicode/utf16\"; func main() { fmt.Println(utf16.IsSurrogate(0xDC00)) }",
        vec!["true"]
    ),
    utf16_is_surrogate_rejects_bmp => (
        "package main; import \"fmt\"; import \"unicode/utf16\"; func main() { fmt.Println(utf16.IsSurrogate(65)) }",
        vec!["false"]
    ),

    unicode_is_letter_arabic => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsLetter('ب')) }",
        vec!["true"]
    ),
    unicode_is_letter_devanagari => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsLetter('क')) }",
        vec!["true"]
    ),
    unicode_is_letter_thai => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsLetter('ก')) }",
        vec!["true"]
    ),
    unicode_is_letter_underscore_rejected => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsLetter('_')) }",
        vec!["false"]
    ),
    unicode_is_digit_roman_numeral_rejected => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsDigit('Ⅳ')) }",
        vec!["false"]
    ),
    unicode_is_digit_circled => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsDigit('③')) }",
        vec!["false"]
    ),

    unicode_to_lower_ascii => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.ToLower('G'))) }",
        vec!["103"]
    ),
    unicode_to_lower_greek_upper => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.ToLower('Λ'))) }",
        vec!["955"]
    ),
    unicode_to_lower_german_eszett => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.ToLower('ẞ'))) }",
        vec!["223"]
    ),
    unicode_to_upper_german_eszett => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.ToUpper('ß'))) }",
        vec!["7838"]
    ),
    unicode_to_lower_turkish_dotless_i => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.ToLower('I'))) }",
        vec!["105"]
    ),
    unicode_to_upper_turkish_dotted_i => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.ToUpper('i'))) }",
        vec!["73"]
    ),

    unicode_simple_fold_sigma => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.SimpleFold('Σ'))) }",
        vec!["963"]
    ),
    unicode_simple_fold_final_sigma => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.SimpleFold('ς'))) }",
        vec!["963"]
    ),
    unicode_simple_fold_kelvin_sign => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.SimpleFold('K'))) }",
        vec!["107"]
    ),
    unicode_simple_fold_angstrom => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.SimpleFold('Å'))) }",
        vec!["229"]
    ),
    unicode_simple_fold_no_mapping => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.SimpleFold('5'))) }",
        vec!["53"]
    ),

    utf8_valid_rune_ascii => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.ValidRune('A')) }",
        vec!["true"]
    ),
    utf8_valid_rune_three_byte => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.ValidRune('世')) }",
        vec!["true"]
    ),
    utf8_valid_rune_emoji => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.ValidRune('🙂')) }",
        vec!["true"]
    ),
    utf8_valid_rune_surrogate_half_rejected => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.ValidRune(0xD800)) }",
        vec!["false"]
    ),
    utf8_valid_rune_out_of_range => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.ValidRune(0x110000)) }",
        vec!["false"]
    ),

    utf8_rune_len_ascii => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.RuneLen('A')) }",
        vec!["1"]
    ),
    utf8_rune_len_two_byte => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.RuneLen('é')) }",
        vec!["2"]
    ),
    utf8_rune_len_three_byte => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.RuneLen('世')) }",
        vec!["3"]
    ),
    utf8_rune_len_four_byte => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.RuneLen('🙂')) }",
        vec!["4"]
    ),
    utf8_rune_len_invalid_negative => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.RuneLen(-1)) }",
        vec!["-1"]
    ),

    utf8_valid_empty_slice => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.Valid([]byte{})) }",
        vec!["true"]
    ),
    utf8_valid_truncated_multibyte => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.Valid([]byte{0xE4, 0xB8})) }",
        vec!["false"]
    ),
    utf8_encode_rune_two_byte => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { buf := make([]byte, 4); n := utf8.EncodeRune(buf, 'é'); fmt.Println(n); fmt.Println(int(buf[0])) }",
        vec!["2", "195"]
    ),
    utf8_encode_rune_four_byte => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { buf := make([]byte, 4); n := utf8.EncodeRune(buf, '🙂'); fmt.Println(n) }",
        vec!["4"]
    ),
    utf8_decode_rune_in_string => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { r, size := utf8.DecodeRuneInString(\"世go\"); fmt.Println(int(r)); fmt.Println(size) }",
        vec!["19990", "3"]
    ),
    utf8_decode_last_rune_in_string => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { s := \"ab\"; r, size := utf8.DecodeLastRuneInString(s); fmt.Println(int(r)); fmt.Println(size) }",
        vec!["98", "1"]
    ),
}

go_compile_cases! {
    utf16_encode_empty_runes => "package main; import \"unicode/utf16\"; func main() { _ = utf16.Encode([]rune{}) }",
    utf16_decode_empty_units => "package main; import \"unicode/utf16\"; func main() { _ = utf16.Decode([]uint16{}) }",
    utf16_encode_rune_bmp_boundary => "package main; import \"unicode/utf16\"; func main() { _, _ = utf16.EncodeRune(0xFFFF) }",
    utf16_decode_rune_invalid_trailing => "package main; import \"unicode/utf16\"; func main() { _ = utf16.DecodeRune(0xDC00, 65535) }",
    utf16_roundtrip_emoji => "package main; import \"unicode/utf16\"; func main() { u := utf16.Encode([]rune(\"🎉\")); _ = utf16.Decode(u) }",

    unicode_is_letter_hangul => "package main; import \"unicode\"; func main() { _ = unicode.IsLetter('한') }",
    unicode_is_letter_hebrew => "package main; import \"unicode\"; func main() { _ = unicode.IsLetter('א') }",
    unicode_is_digit_devanagari => "package main; import \"unicode\"; func main() { _ = unicode.IsDigit('४') }",
    unicode_is_digit_thai => "package main; import \"unicode\"; func main() { _ = unicode.IsDigit('๙') }",
    unicode_to_lower_hangul => "package main; import \"unicode\"; func main() { _ = unicode.ToLower('A') }",
    unicode_to_upper_cyrillic => "package main; import \"unicode\"; func main() { _ = unicode.ToUpper('ж') }",
    unicode_simple_fold_long_s => "package main; import \"unicode\"; func main() { _ = unicode.SimpleFold('ſ') }",
    unicode_simple_fold_micro_sign => "package main; import \"unicode\"; func main() { _ = unicode.SimpleFold('µ') }",

    utf8_valid_rune_max => "package main; import \"unicode/utf8\"; func main() { _ = utf8.ValidRune(0x10FFFF) }",
    utf8_rune_len_out_of_range => "package main; import \"unicode/utf8\"; func main() { _ = utf8.RuneLen(0x110000) }",
    utf8_encode_rune_to_string => "package main; import \"unicode/utf8\"; func main() { _ = utf8.EncodeRuneToString('€') }",
    utf8_decode_rune_in_string_empty => "package main; import \"unicode/utf8\"; func main() { _, _ = utf8.DecodeRuneInString(\"\") }",
    utf8_decode_last_rune_in_string_unicode => "package main; import \"unicode/utf8\"; func main() { _, _ = utf8.DecodeLastRuneInString(\"日\") }",
    utf8_append_rune_existing_buffer => "package main; import \"unicode/utf8\"; func main() { _ = utf8.AppendRune([]byte(\"a\"), 'b') }",
    utf8_full_rune_at_offset => "package main; import \"unicode/utf8\"; func main() { _ = utf8.FullRuneAt([]byte(\"世\"), 0) }",
    utf8_full_rune_in_string_at => "package main; import \"unicode/utf8\"; func main() { _ = utf8.FullRuneInStringAt(\"a世\", 1) }",
    utf8_valid_string_empty => "package main; import \"unicode/utf8\"; func main() { _ = utf8.ValidString(\"\") }",
    utf8_valid_string_emoji => "package main; import \"unicode/utf8\"; func main() { _ = utf8.ValidString(\"🙂\") }",
    utf8_rune_count_empty => "package main; import \"unicode/utf8\"; func main() { _ = utf8.RuneCount([]byte{}) }",
    utf8_rune_count_in_string_emoji => "package main; import \"unicode/utf8\"; func main() { _ = utf8.RuneCountInString(\"🙂\") }",
}
