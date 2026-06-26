//! unicode/utf8 package: Valid, ValidString, RuneCount, EncodeRune, DecodeRune,
//! FullRune, and UTF-8 string iteration — plus distinct rune literal semantics.

use crate::helpers::*;

go_run_cases! {
    utf8_valid_ascii_bytes => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.Valid([]byte(\"go\"))) }",
        vec!["true"]
    ),
    utf8_valid_multibyte_bytes => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.Valid([]byte(\"日本\"))) }",
        vec!["true"]
    ),
    utf8_valid_rejects_lone_continuation => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.Valid([]byte{0x80})) }",
        vec!["false"]
    ),
    utf8_valid_rejects_overlong_encoding => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.Valid([]byte{0xC0, 0xAF})) }",
        vec!["false"]
    ),
    utf8_valid_string_ascii => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.ValidString(\"hello\")) }",
        vec!["true"]
    ),
    utf8_valid_string_with_combining_mark => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.ValidString(\"café\")) }",
        vec!["true"]
    ),
    utf8_rune_count_ascii_bytes => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.RuneCount([]byte(\"hello\"))) }",
        vec!["5"]
    ),
    utf8_rune_count_multibyte_bytes => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.RuneCount([]byte(\"日本\"))) }",
        vec!["2"]
    ),
    utf8_rune_count_in_string_accented => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.RuneCountInString(\"héllo\")) }",
        vec!["5"]
    ),
    utf8_len_differs_from_rune_count => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { s := \"日本\"; fmt.Println(len(s)); fmt.Println(utf8.RuneCountInString(s)) }",
        vec!["6", "2"]
    ),
    utf8_encode_rune_ascii => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { buf := make([]byte, 4); n := utf8.EncodeRune(buf, 'A'); fmt.Println(n); fmt.Println(int(buf[0])) }",
        vec!["1", "65"]
    ),
    utf8_encode_rune_three_byte => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { buf := make([]byte, 4); n := utf8.EncodeRune(buf, '世'); fmt.Println(n); fmt.Println(int(buf[0])); fmt.Println(int(buf[1])); fmt.Println(int(buf[2])) }",
        vec!["3", "228", "184", "150"]
    ),
    utf8_decode_rune_ascii => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { r, size := utf8.DecodeRune([]byte(\"Z\")); fmt.Println(int(r)); fmt.Println(size) }",
        vec!["90", "1"]
    ),
    utf8_decode_rune_multibyte => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { r, size := utf8.DecodeRune([]byte(\"世\")); fmt.Println(int(r)); fmt.Println(size) }",
        vec!["19990", "3"]
    ),
    utf8_decode_rune_invalid_byte => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { r, size := utf8.DecodeRune([]byte{0xff}); fmt.Println(int(r)); fmt.Println(size) }",
        vec!["65533", "1"]
    ),
    utf8_full_rune_complete_ascii => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.FullRune([]byte(\"x\"))) }",
        vec!["true"]
    ),
    utf8_full_rune_incomplete_lead_byte => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.FullRune([]byte{0xE4})) }",
        vec!["false"]
    ),
    utf8_full_rune_in_string_leading => (
        "package main; import \"fmt\"; import \"unicode/utf8\"; func main() { fmt.Println(utf8.FullRuneInString(\"界\")) }",
        vec!["true"]
    ),
    utf8_string_range_byte_indices => (
        "package main; import \"fmt\"; func main() { first, second := -1, -1; step := 0; for i, _ := range \"a世\" { if step == 0 { first = i }; if step == 1 { second = i }; step++ }; fmt.Println(first); fmt.Println(second) }",
        vec!["0", "1"]
    ),
    utf8_string_range_rune_values => (
        "package main; import \"fmt\"; func main() { total := 0; for _, r := range \"日本語\" { total += int(r) }; fmt.Println(total) }",
        vec!["87983"]
    ),
    rune_literal_escape_newline => (
        "package main; import \"fmt\"; func main() { fmt.Println(int('\\n')) }",
        vec!["10"]
    ),
    rune_literal_unicode_short_escape => (
        "package main; import \"fmt\"; func main() { fmt.Println(int('\\u03BB')) }",
        vec!["955"]
    ),
    rune_literal_unicode_long_escape => (
        "package main; import \"fmt\"; func main() { fmt.Println(int('\\U00000041')) }",
        vec!["65"]
    ),
    rune_literal_backslash_escape => (
        "package main; import \"fmt\"; func main() { fmt.Println(int('\\\\')) }",
        vec!["92"]
    ),
    rune_literal_single_quote_escape => (
        "package main; import \"fmt\"; func main() { fmt.Println(int('\\'')) }",
        vec!["39"]
    ),
    rune_literal_const_and_compare => (
        "package main; import \"fmt\"; const letter rune = 'λ'; func main() { fmt.Println(int(letter)); fmt.Println(letter == '\\u03BB') }",
        vec!["955", "true"]
    ),
    rune_literal_arithmetic => (
        "package main; import \"fmt\"; func main() { fmt.Println(int('A' + 1)) }",
        vec!["66"]
    ),
}

go_compile_cases! {
    utf8_decode_rune_in_string_compile => "package main; import \"unicode/utf8\"; func main() { _, _ = utf8.DecodeRuneInString(\"世\") }",
    utf8_encode_rune_to_string_compile => "package main; import \"unicode/utf8\"; func main() { _ = utf8.EncodeRuneToString('€') }",
    utf8_valid_rune_compile => "package main; import \"unicode/utf8\"; func main() { _ = utf8.ValidRune('🙂') }",
    utf8_rune_len_compile => "package main; import \"unicode/utf8\"; func main() { _ = utf8.RuneLen('世') }",
    utf8_append_rune_compile => "package main; import \"unicode/utf8\"; func main() { buf := []byte{}; _ = utf8.AppendRune(buf, 'a') }",
    rune_literal_tab_escape_compile => "package main; func main() { _ = '\\t' }",
    rune_literal_emoji_compile => "package main; func main() { _ = '🙂' }",
    rune_slice_unicode_compile => "package main; func main() { rs := []rune(\"café\"); _ = rs[4] }",
}
