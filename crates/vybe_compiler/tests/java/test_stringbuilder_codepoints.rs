use crate::helpers::run_main;

#[test]
fn stringbuilder_code_point_count_ascii_full_range() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("hello"); System.out.println(sb.codePointCount(0, 5));"#);
    assert_eq!(out, vec!["5"]);
}

#[test]
fn stringbuilder_code_point_count_ascii_partial_range() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("abcdef"); System.out.println(sb.codePointCount(1, 4));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringbuilder_code_point_count_empty_range() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("abc"); System.out.println(sb.codePointCount(1, 1));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stringbuilder_code_point_count_single_char_range() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("xyz"); System.out.println(sb.codePointCount(0, 1));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn stringbuilder_code_point_count_supplementary_single_glyph() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("a\uD83D\uDE00b"); System.out.println(sb.codePointCount(0, 4));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringbuilder_code_point_count_supplementary_only() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("\uD83D\uDE00"); System.out.println(sb.codePointCount(0, 2));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn stringbuilder_code_point_count_after_append() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("ab"); sb.append("cd"); System.out.println(sb.codePointCount(0, 4));"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn stringbuilder_code_point_count_after_insert() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("ac"); sb.insert(1, "b"); System.out.println(sb.codePointCount(0, 3));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringbuilder_code_point_count_after_delete() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("abcde"); sb.delete(2, 4); System.out.println(sb.codePointCount(0, sb.length()));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringbuilder_code_point_count_after_replace() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("aXc"); sb.replace(1, 2, "BB"); System.out.println(sb.codePointCount(0, 4));"#);
    assert_eq!(out, vec!["4"]);
}

#[test]
fn character_to_code_point_from_high_low_surrogates() {
    let out = run_main(r#"System.out.println(Character.toCodePoint('\uD83D', '\uDE00'));"#);
    assert_eq!(out, vec!["128512"]);
}

#[test]
fn character_to_code_point_basic_latin_letter() {
    let out = run_main(r#"System.out.println(Character.toCodePoint('A', 'B'));"#);
    assert_eq!(out, vec!["65"]);
}

#[test]
fn character_to_code_point_second_arg_ignored_for_bmp() {
    let out = run_main(r#"System.out.println(Character.toCodePoint('Z', '\u0000'));"#);
    assert_eq!(out, vec!["90"]);
}

#[test]
fn character_to_code_point_smiley_emoji_code() {
    let out = run_main(r#"int cp = Character.toCodePoint('\uD83D', '\uDE04'); System.out.println(cp);"#);
    assert_eq!(out, vec!["128516"]);
}

#[test]
fn character_to_code_point_from_supplementary_pair_in_builder() {
    let out = run_main(r#"int cp = Character.toCodePoint('\uD83C', '\uDF89'); StringBuilder sb = new StringBuilder(); sb.appendCodePoint(cp); System.out.println(sb.codePointCount(0, 2));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn stringbuilder_code_point_count_unicode_letter_with_diacritic() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("\u00e9"); System.out.println(sb.codePointCount(0, 1));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn stringbuilder_code_point_count_mixed_ascii_and_supplementary() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("hi\uD83D\uDE00"); System.out.println(sb.codePointCount(0, 4));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringbuilder_code_point_count_from_offset_one() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("x\uD83D\uDE00y"); System.out.println(sb.codePointCount(1, 4));"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stringbuilder_code_point_count_two_supplementary_chars() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("\uD83D\uDE00\uD83D\uDE04"); System.out.println(sb.codePointCount(0, 4));"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stringbuilder_code_point_count_after_set_length_truncation() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("abcdef"); sb.setLength(3); System.out.println(sb.codePointCount(0, 3));"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn stringbuilder_code_point_count_on_empty_builder() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder(); System.out.println(sb.codePointCount(0, 0));"#);
    assert_eq!(out, vec!["0"]);
}

#[test]
fn stringbuilder_code_point_count_cjk_character() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("\u4F60\u597D"); System.out.println(sb.codePointCount(0, 2));"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn character_to_code_point_matches_append_code_point() {
    let out = run_main(r#"int cp = Character.toCodePoint('\uD83D', '\uDE00'); StringBuilder sb = new StringBuilder(); sb.appendCodePoint(cp); System.out.println(sb.length());"#);
    assert_eq!(out, vec!["2"]);
}

#[test]
fn stringbuilder_code_point_count_subrange_excluding_supplementary() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("a\uD83D\uDE00b"); System.out.println(sb.codePointCount(0, 1));"#);
    assert_eq!(out, vec!["1"]);
}

#[test]
fn stringbuilder_code_point_count_equals_length_for_pure_ascii() {
    let out = run_main(r#"StringBuilder sb = new StringBuilder("vybe"); System.out.println(sb.codePointCount(0, sb.length()) == sb.length());"#);
    assert_eq!(out, vec!["true"]);
}
