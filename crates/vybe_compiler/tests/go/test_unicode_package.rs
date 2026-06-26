//! unicode package: IsLetter, IsDigit, ToUpper, and In range-table membership.
//!
//! Distinct from `test_unicode_utf8.rs`, which covers UTF-8 encoding/decoding only.

use crate::helpers::*;

go_run_cases! {
    unicode_is_letter_ascii_upper => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsLetter('A')) }",
        vec!["true"]
    ),
    unicode_is_letter_ascii_lower => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsLetter('z')) }",
        vec!["true"]
    ),
    unicode_is_letter_greek_alpha => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsLetter('α')) }",
        vec!["true"]
    ),
    unicode_is_letter_cyrillic => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsLetter('Ж')) }",
        vec!["true"]
    ),
    unicode_is_letter_han => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsLetter('世')) }",
        vec!["true"]
    ),
    unicode_is_letter_rejects_digit => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsLetter('7')) }",
        vec!["false"]
    ),
    unicode_is_letter_rejects_space => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsLetter(' ')) }",
        vec!["false"]
    ),
    unicode_is_letter_rejects_punct => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsLetter('!')) }",
        vec!["false"]
    ),
    unicode_is_digit_zero => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsDigit('0')) }",
        vec!["true"]
    ),
    unicode_is_digit_nine => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsDigit('9')) }",
        vec!["true"]
    ),
    unicode_is_digit_fullwidth => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsDigit('５')) }",
        vec!["true"]
    ),
    unicode_is_digit_arabic_indic => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsDigit('٠')) }",
        vec!["true"]
    ),
    unicode_is_digit_rejects_letter => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsDigit('A')) }",
        vec!["false"]
    ),
    unicode_is_digit_rejects_superscript => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.IsDigit('²')) }",
        vec!["false"]
    ),
    unicode_to_upper_ascii_lower => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.ToUpper('a'))) }",
        vec!["65"]
    ),
    unicode_to_upper_ascii_already_upper => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.ToUpper('Z'))) }",
        vec!["90"]
    ),
    unicode_to_upper_greek_lower => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.ToUpper('λ'))) }",
        vec!["923"]
    ),
    unicode_to_upper_digit_unchanged => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.ToUpper('5'))) }",
        vec!["53"]
    ),
    unicode_to_upper_punct_unchanged => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(int(unicode.ToUpper('!'))) }",
        vec!["33"]
    ),
    unicode_in_greek_alpha => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.In('α', unicode.Greek)) }",
        vec!["true"]
    ),
    unicode_in_latin_capital => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.In('A', unicode.Latin)) }",
        vec!["true"]
    ),
    unicode_in_rejects_latin_for_greek => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.In('α', unicode.Latin)) }",
        vec!["false"]
    ),
    unicode_in_rejects_greek_for_latin => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.In('A', unicode.Greek)) }",
        vec!["false"]
    ),
    unicode_in_digit_table => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.In('5', unicode.Digit)) }",
        vec!["true"]
    ),
    unicode_in_letter_table_han => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.In('世', unicode.Letter)) }",
        vec!["true"]
    ),
    unicode_in_han_script => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.In('世', unicode.Han)) }",
        vec!["true"]
    ),
    unicode_in_punct_exclaim => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.In('!', unicode.Punct)) }",
        vec!["true"]
    ),
    unicode_in_cyrillic => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.In('Ж', unicode.Cyrillic)) }",
        vec!["true"]
    ),
    unicode_in_multiple_range_tables => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.In('5', unicode.Digit, unicode.Number)) }",
        vec!["true"]
    ),
    unicode_in_space_table => (
        "package main; import \"fmt\"; import \"unicode\"; func main() { fmt.Println(unicode.In(' ', unicode.Space)) }",
        vec!["true"]
    ),
}

go_compile_cases! {
    unicode_is_upper_compile => "package main; import \"unicode\"; func main() { _ = unicode.IsUpper('A') }",
    unicode_is_lower_compile => "package main; import \"unicode\"; func main() { _ = unicode.IsLower('a') }",
    unicode_to_lower_compile => "package main; import \"unicode\"; func main() { _ = unicode.ToLower('A') }",
    unicode_is_space_compile => "package main; import \"unicode\"; func main() { _ = unicode.IsSpace('\\t') }",
    unicode_is_number_compile => "package main; import \"unicode\"; func main() { _ = unicode.IsNumber('²') }",
    unicode_in_upper_table_compile => "package main; import \"unicode\"; func main() { _ = unicode.In('A', unicode.Upper) }",
}
