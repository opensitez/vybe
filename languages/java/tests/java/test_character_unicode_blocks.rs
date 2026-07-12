use crate::helpers::run_main;

#[test]
fn character_unicode_block_of_uppercase_a_is_basic_latin() {
    let out = run_main("System.out.println(Character.UnicodeBlock.of('A').toString());");
    assert_eq!(out, vec!["BASIC_LATIN"]);
}

#[test]
fn character_unicode_block_of_digit_zero_is_basic_latin() {
    let out = run_main("System.out.println(Character.UnicodeBlock.of('0').toString());");
    assert_eq!(out, vec!["BASIC_LATIN"]);
}

#[test]
fn character_unicode_block_of_greek_alpha() {
    let out = run_main("System.out.println(Character.UnicodeBlock.of('\\u03B1').toString());");
    assert_eq!(out, vec!["GREEK"]);
}

#[test]
fn character_unicode_block_of_cyrillic_letter() {
    let out = run_main("System.out.println(Character.UnicodeBlock.of('\\u0416').toString());");
    assert_eq!(out, vec!["CYRILLIC"]);
}

#[test]
fn character_unicode_block_of_hiragana_char() {
    let out = run_main("System.out.println(Character.UnicodeBlock.of('\\u3042').toString());");
    assert_eq!(out, vec!["HIRAGANA"]);
}

#[test]
fn character_get_type_uppercase_letter() {
    let out = run_main("System.out.println(Character.getType('Z'));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn character_get_type_lowercase_letter() {
    let out = run_main("System.out.println(Character.getType('z'));");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn character_get_type_decimal_digit() {
    let out = run_main("System.out.println(Character.getType('7'));");
    assert_eq!(out, vec!["9"]);
}

#[test]
fn character_get_type_space_separator() {
    let out = run_main("System.out.println(Character.getType(' '));");
    assert_eq!(out, vec!["12"]);
}

#[test]
fn character_get_type_currency_symbol() {
    let out = run_main("System.out.println(Character.getType('$'));");
    assert_eq!(out, vec!["26"]);
}

#[test]
fn character_is_java_identifier_start_accepts_letter() {
    let out = run_main("System.out.println(Character.isJavaIdentifierStart('j'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_java_identifier_start_accepts_dollar() {
    let out = run_main("System.out.println(Character.isJavaIdentifierStart('$'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_java_identifier_start_accepts_underscore() {
    let out = run_main("System.out.println(Character.isJavaIdentifierStart('_'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_java_identifier_start_rejects_digit() {
    let out = run_main("System.out.println(Character.isJavaIdentifierStart('9'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_java_identifier_part_accepts_digit() {
    let out = run_main("System.out.println(Character.isJavaIdentifierPart('4'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_java_identifier_part_accepts_letter() {
    let out = run_main("System.out.println(Character.isJavaIdentifierPart('k'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_java_identifier_part_rejects_space() {
    let out = run_main("System.out.println(Character.isJavaIdentifierPart(' '));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_unicode_identifier_start_accepts_letter() {
    let out = run_main("System.out.println(Character.isUnicodeIdentifierStart('A'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_unicode_identifier_start_rejects_punctuation() {
    let out = run_main("System.out.println(Character.isUnicodeIdentifierStart('!'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_unicode_identifier_part_accepts_connector() {
    let out = run_main("System.out.println(Character.isUnicodeIdentifierPart('_'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_unicode_identifier_part_accepts_digit() {
    let out = run_main("System.out.println(Character.isUnicodeIdentifierPart('3'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_iso_control_on_tab() {
    let out = run_main("System.out.println(Character.isISOControl('\\t'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_iso_control_on_newline() {
    let out = run_main("System.out.println(Character.isISOControl('\\n'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_iso_control_rejects_printable() {
    let out = run_main("System.out.println(Character.isISOControl('A'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_mirrored_on_open_paren() {
    let out = run_main("System.out.println(Character.isMirrored('('));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_mirrored_on_close_bracket() {
    let out = run_main("System.out.println(Character.isMirrored(']'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_mirrored_rejects_letter() {
    let out = run_main("System.out.println(Character.isMirrored('x'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_to_title_case_from_lowercase_i() {
    let out = run_main("System.out.println(Character.toTitleCase('i'));");
    assert_eq!(out, vec!["I"]);
}

#[test]
fn character_to_title_case_leaves_uppercase() {
    let out = run_main("System.out.println(Character.toTitleCase('M'));");
    assert_eq!(out, vec!["M"]);
}

#[test]
fn character_is_defined_on_ascii_letter() {
    let out = run_main("System.out.println(Character.isDefined('Q'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_defined_rejects_unassigned_codepoint() {
    let out = run_main("System.out.println(Character.isDefined('\\uFFFF'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_bmp_code_point_on_basic_latin() {
    let out = run_main("System.out.println(Character.isBmpCodePoint((int)'A'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_bmp_code_point_rejects_supplementary() {
    let out = run_main("System.out.println(Character.isBmpCodePoint(0x1F600));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_valid_code_point_on_max_bmp() {
    let out = run_main("System.out.println(Character.isValidCodePoint(0xFFFF));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_valid_code_point_rejects_too_large() {
    let out = run_main("System.out.println(Character.isValidCodePoint(0x110000));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_high_surrogate_on_leading_unit() {
    let out = run_main("System.out.println(Character.isHighSurrogate('\\uD800'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_high_surrogate_rejects_low_surrogate() {
    let out = run_main("System.out.println(Character.isHighSurrogate('\\uDC00'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_low_surrogate_on_trailing_unit() {
    let out = run_main("System.out.println(Character.isLowSurrogate('\\uDC00'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_low_surrogate_rejects_high_surrogate() {
    let out = run_main("System.out.println(Character.isLowSurrogate('\\uD800'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_is_surrogate_pair_on_emoji() {
    let out = run_main("System.out.println(Character.isSurrogatePair('\\uD83D', '\\uDE00'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_is_surrogate_pair_rejects_two_high() {
    let out = run_main("System.out.println(Character.isSurrogatePair('\\uD800', '\\uD801'));");
    assert_eq!(out, vec!["false"]);
}

#[test]
fn character_code_point_count_ascii_string() {
    let out = run_main("System.out.println(Character.codePointCount(\"hello\", 0, 5));");
    assert_eq!(out, vec!["5"]);
}

#[test]
fn character_code_point_count_with_emoji() {
    let out = run_main("System.out.println(Character.codePointCount(\"a\\uD83D\\uDE00b\", 0, 4));");
    assert_eq!(out, vec!["3"]);
}

#[test]
fn character_to_code_point_from_surrogate_pair() {
    let out = run_main("System.out.println(Character.toCodePoint('\\uD83D', '\\uDE00'));");
    assert_eq!(out, vec!["128512"]);
}

#[test]
fn character_char_count_from_code_point_ascii() {
    let out = run_main("System.out.println(Character.charCount((int)'A'));");
    assert_eq!(out, vec!["1"]);
}

#[test]
fn character_char_count_from_code_point_emoji() {
    let out = run_main("System.out.println(Character.charCount(0x1F600));");
    assert_eq!(out, vec!["2"]);
}

#[test]
fn character_unicode_block_of_box_drawing() {
    let out = run_main("System.out.println(Character.UnicodeBlock.of('\\u2500').toString());");
    assert_eq!(out, vec!["BOX_DRAWING"]);
}

#[test]
fn character_get_type_line_separator() {
    let out = run_main("System.out.println(Character.getType('\\u2028'));");
    assert_eq!(out, vec!["13"]);
}

#[test]
fn character_is_java_identifier_part_accepts_dollar() {
    let out = run_main("System.out.println(Character.isJavaIdentifierPart('$'));");
    assert_eq!(out, vec!["true"]);
}

#[test]
fn character_code_point_at_in_string_for_emoji() {
    let out = run_main("System.out.println(Character.codePointAt(\"x\\uD83D\\uDE00y\", 1));");
    assert_eq!(out, vec!["128512"]);
}
