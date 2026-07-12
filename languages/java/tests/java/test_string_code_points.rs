use crate::helpers::run_main;

#[test]
fn character_to_chars_single_bmp() {
    let out = run_main(r#"System.out.println(new String(Character.toChars(65)));"#);
    assert_eq!(out, vec!["A"]);
}

#[test]
fn character_to_chars_supplementary() {
    let out = run_main(r#"System.out.println(new String(Character.toChars(0x1F600)).length());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn character_to_code_point_surrogate_pair() {
    let out = run_main(r#"System.out.println(Character.toCodePoint('\uD83D', '\uDE00'));"#);
    assert_eq!(out, vec!["128512"]);
}

#[test]
fn character_is_surrogate_high() {
    let out = run_main(r#"System.out.println(Character.isSurrogate('\uD800'));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_surrogate_low() {
    let out = run_main(r#"System.out.println(Character.isSurrogate('\uDFFF'));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_surrogate_bmp_false() {
    let out = run_main(r#"System.out.println(Character.isSurrogate('A'));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_high_surrogate_true() {
    let out = run_main(r#"System.out.println(Character.isHighSurrogate('\uD83D'));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_high_surrogate_false() {
    let out = run_main(r#"System.out.println(Character.isHighSurrogate('\uDE00'));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_low_surrogate_true() {
    let out = run_main(r#"System.out.println(Character.isLowSurrogate('\uDE00'));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_low_surrogate_false() {
    let out = run_main(r#"System.out.println(Character.isLowSurrogate('\uD83D'));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_high_surrogate_of_emoji() {
    let out = run_main(
        r#"System.out.println(Character.isHighSurrogate(Character.highSurrogate(0x1F600)));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_low_surrogate_of_emoji() {
    let out = run_main(
        r#"System.out.println(Character.isLowSurrogate(Character.lowSurrogate(0x1F600)));"#,
    );
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_char_count_bmp() {
    let out = run_main(r#"System.out.println(Character.charCount(65));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn character_char_count_supplementary() {
    let out = run_main(r#"System.out.println(Character.charCount(0x1F600));"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn character_code_point_at_start() {
    let out = run_main(r#"System.out.println(Character.codePointAt("Abc", 0));"#);
    assert_eq!(out, vec!["65"]);
}

#[test]
fn character_code_point_before_end() {
    let out = run_main(r#"System.out.println(Character.codePointBefore("Abc", 3));"#);
    assert_eq!(out, vec!["99"]);
}

#[test]
fn character_code_point_count_ascii() {
    let out = run_main(r#"System.out.println(Character.codePointCount("hello", 0, 5));"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn character_code_point_count_emoji() {
    let out = run_main(r#"System.out.println(Character.codePointCount("a\uD83D\uDE00b", 0, 4));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn character_offset_by_code_points_forward() {
    let out = run_main(r#"System.out.println(Character.offsetByCodePoints("abcd", 0, 2));"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn character_offset_by_code_points_backward() {
    let out = run_main(r#"System.out.println(Character.offsetByCodePoints("abcd", 4, -2));"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn string_new_from_code_point_array() {
    let out = run_main(r#"int[] cps = {72, 105}; System.out.println(new String(cps, 0, 2));"#);
    assert_eq!(out, vec!["Hi"]);
}

#[test]
fn string_code_points_count_stream() {
    let out = run_main(r#"String s = "abc"; System.out.println(s.codePoints().count());"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn string_code_points_sum_ascii() {
    let out = run_main(r#"String s = "AB"; System.out.println(s.codePoints().sum());"#);
    assert_eq!(out, vec!["131"]);
}

#[test]
fn string_chars_count() {
    let out = run_main(r#"String s = "abc"; System.out.println(s.chars().count());"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn string_chars_sum() {
    let out = run_main(r#"String s = "AB"; System.out.println(s.chars().sum());"#);
    assert_eq!(out, vec!["131"]);
}

#[test]
fn string_code_point_at_bmp() {
    let out = run_main(r#"String s = "Z"; System.out.println(s.codePointAt(0));"#);
    assert_eq!(out, vec!["90"]);
}

#[test]
fn string_code_point_before_end() {
    let out = run_main(r#"String s = "ab"; System.out.println(s.codePointBefore(2));"#);
    assert_eq!(out, vec!["98"]);
}

#[test]
fn string_offset_by_code_points_positive() {
    let out = run_main(r#"String s = "abcd"; System.out.println(s.offsetByCodePoints(1, 2));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn string_offset_by_code_points_negative() {
    let out = run_main(r#"String s = "abcd"; System.out.println(s.offsetByCodePoints(3, -2));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn string_code_point_count_range() {
    let out = run_main(r#"String s = "hello"; System.out.println(s.codePointCount(1, 4));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn character_is_valid_code_point_ascii() {
    let out = run_main(r#"System.out.println(Character.isValidCodePoint(65));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_valid_code_point_max() {
    let out = run_main(r#"System.out.println(Character.isValidCodePoint(0x10FFFF));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_valid_code_point_too_high() {
    let out = run_main(r#"System.out.println(Character.isValidCodePoint(0x110000));"#);
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_bmp_code_point() {
    let out = run_main(r#"System.out.println(Character.isBmpCodePoint(0xFFFF));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_supplementary_code_point() {
    let out = run_main(r#"System.out.println(Character.isSupplementaryCodePoint(0x1F600));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_to_chars_two_element() {
    let out =
        run_main(r#"char[] arr = Character.toChars(0x1F600); System.out.println(arr.length);"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn character_digit_parse_hex() {
    let out = run_main(r#"System.out.println(Character.digit('A', 16));"#);
    assert_eq!(out, vec!["10"]);
}

#[test]
fn character_digit_parse_decimal() {
    let out = run_main(r#"System.out.println(Character.digit('7', 10));"#);
    assert_eq!(out, vec!["7"]);
}

#[test]
fn character_for_digit_hex() {
    let out = run_main(r#"System.out.println(Character.forDigit(10, 16));"#);
    assert_eq!(out, vec!["a"]);
}

#[test]
fn character_for_digit_decimal() {
    let out = run_main(r#"System.out.println(Character.forDigit(5, 10));"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn string_code_points_max_ascii() {
    let out = run_main(r#"String s = "z"; System.out.println(s.codePoints().max().getAsInt());"#);
    assert_eq!(out, vec!["122"]);
}

#[test]
fn string_code_points_min_ascii() {
    let out = run_main(r#"String s = "a"; System.out.println(s.codePoints().min().getAsInt());"#);
    assert_eq!(out, vec!["97"]);
}

#[test]
fn string_chars_max() {
    let out = run_main(r#"String s = "Z"; System.out.println(s.chars().max().getAsInt());"#);
    assert_eq!(out, vec!["90"]);
}

#[test]
fn string_code_point_at_supplementary() {
    let out = run_main(r#"String s = "\uD83D\uDE00"; System.out.println(s.codePointAt(0));"#);
    assert_eq!(out, vec!["128512"]);
}

#[test]
fn string_code_point_before_supplementary() {
    let out = run_main(r#"String s = "x\uD83D\uDE00"; System.out.println(s.codePointBefore(4));"#);
    assert_eq!(out, vec!["128512"]);
}

#[test]
fn character_compare_ascii() {
    let out = run_main(r#"System.out.println(Character.compare('a', 'b') < 0);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_reverse_bytes() {
    let out = run_main(r#"System.out.println((int) Character.reverseBytes((char) 0x1234));"#);
    assert_eq!(out, vec!["13330"]);
}

#[test]
fn string_code_points_distinct_count() {
    let out =
        run_main(r#"String s = "aab"; System.out.println(s.codePoints().distinct().count());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn string_chars_distinct_count() {
    let out = run_main(r#"String s = "aab"; System.out.println(s.chars().distinct().count());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn string_offset_by_code_points_emoji() {
    let out =
        run_main(r#"String s = "a\uD83D\uDE00c"; System.out.println(s.offsetByCodePoints(0, 2));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn character_is_defined_ascii() {
    let out = run_main(r#"System.out.println(Character.isDefined('A'));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_defined_emoji() {
    let out = run_main(r#"System.out.println(Character.isDefined(0x1F600));"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_get_type_letter() {
    let out =
        run_main(r#"System.out.println(Character.getType('A') == Character.UPPERCASE_LETTER);"#);
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_get_numeric_value_digit() {
    let out = run_main(r#"System.out.println(Character.getNumericValue('9'));"#);
    assert_eq!(out, vec!["9"]);
}
