use crate::helpers::run_main;

#[test]
fn matcher_replace_first_substitutes_only_leading_digit() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d"); java.util.regex.Matcher m = p.matcher("a1b2c"); System.out.println(m.replaceFirst("#"));"##,
    );
    assert_eq!(out, vec!["a#b2c"]);
}

#[test]
fn matcher_replace_first_leaves_second_match() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d+"); java.util.regex.Matcher m = p.matcher("9aa99bb"); System.out.println(m.replaceFirst("X"));"##,
    );
    assert_eq!(out, vec!["Xaa99bb"]);
}

#[test]
fn matcher_replace_first_on_no_match_returns_original() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d+"); java.util.regex.Matcher m = p.matcher("letters"); System.out.println(m.replaceFirst("0"));"##,
    );
    assert_eq!(out, vec!["letters"]);
}

#[test]
fn matcher_replace_first_with_word_pattern() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\w+"); java.util.regex.Matcher m = p.matcher("hi there"); System.out.println(m.replaceFirst("[word]"));"##,
    );
    assert_eq!(out, vec!["[word] there"]);
}

#[test]
fn matcher_replace_first_backreference_reorders_groups() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("(\\w+)-(\w+)"); java.util.regex.Matcher m = p.matcher("foo-bar-baz"); System.out.println(m.replaceFirst("$2_$1"));"##,
    );
    assert_eq!(out, vec!["bar_foo-baz"]);
}

#[test]
fn matcher_replace_first_backreference_swaps_pair() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("(a)(b)"); java.util.regex.Matcher m = p.matcher("abcc"); System.out.println(m.replaceFirst("$2$1"));"##,
    );
    assert_eq!(out, vec!["bacc"]);
}

#[test]
fn matcher_replace_all_changes_every_digit_not_first_only() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d"); java.util.regex.Matcher m = p.matcher("a1b2c"); System.out.println(m.replaceAll("#"));"##,
    );
    assert_eq!(out, vec!["a#b#c"]);
}

#[test]
fn matcher_replace_first_vs_replace_all_on_two_digits() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d"); java.util.regex.Matcher m1 = p.matcher("x1y2"); java.util.regex.Matcher m2 = p.matcher("x1y2"); System.out.println(m1.replaceFirst("*")); System.out.println(m2.replaceAll("*"));"##,
    );
    assert_eq!(out, vec!["x*1y2", "x*y*"]);
}

#[test]
fn matcher_replace_first_strips_leading_zeros_once() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("^0+"); java.util.regex.Matcher m = p.matcher("007007"); System.out.println(m.replaceFirst(""));"##,
    );
    assert_eq!(out, vec!["7007"]);
}

#[test]
fn matcher_replace_all_strips_all_zero_runs() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("0+"); java.util.regex.Matcher m = p.matcher("a00b000c"); System.out.println(m.replaceAll(""));"##,
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn matcher_append_replacement_wraps_each_word() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\w+"); java.util.regex.Matcher m = p.matcher("hi there"); StringBuffer sb = new StringBuffer(); while (m.find()) { m.appendReplacement(sb, "[$0]"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["[hi] [there]"]);
}

#[test]
fn matcher_append_replacement_with_group_backreference() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("(\\w)(\w)"); java.util.regex.Matcher m = p.matcher("ab"); StringBuffer sb = new StringBuffer(); if (m.find()) { m.appendReplacement(sb, "$2$1"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["ba"]);
}

#[test]
fn matcher_append_tail_appends_unmatched_suffix() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d+"); java.util.regex.Matcher m = p.matcher("num42end"); StringBuffer sb = new StringBuffer(); if (m.find()) { m.appendReplacement(sb, "#"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["num#end"]);
}

#[test]
fn matcher_append_tail_on_no_match_copies_whole_input() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d+"); java.util.regex.Matcher m = p.matcher("text"); StringBuffer sb = new StringBuffer(); m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["text"]);
}

#[test]
fn matcher_manual_loop_replaces_two_digit_runs() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d+"); java.util.regex.Matcher m = p.matcher("a12b34"); StringBuffer sb = new StringBuffer(); while (m.find()) { m.appendReplacement(sb, "N"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["aNbN"]);
}

#[test]
fn matcher_manual_loop_preserves_prefix_before_first_match() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("cat"); java.util.regex.Matcher m = p.matcher("copycat"); StringBuffer sb = new StringBuffer(); if (m.find()) { m.appendReplacement(sb, "dog"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["copydog"]);
}

#[test]
fn matcher_append_replacement_literal_dollar_sign() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a"); java.util.regex.Matcher m = p.matcher("aba"); StringBuffer sb = new StringBuffer(); while (m.find()) { m.appendReplacement(sb, "\\$"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["$b$"]);
}

#[test]
fn matcher_replace_first_on_email_local_part() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("[a-z]+"); java.util.regex.Matcher m = p.matcher("user@host"); System.out.println(m.replaceFirst("name"));"##,
    );
    assert_eq!(out, vec!["name@host"]);
}

#[test]
fn matcher_replace_first_on_hyphenated_token() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("-"); java.util.regex.Matcher m = p.matcher("a-b-c"); System.out.println(m.replaceFirst("_"));"##,
    );
    assert_eq!(out, vec!["a_b-c"]);
}

#[test]
fn matcher_replace_first_uppercases_first_vowel() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("[aeiou]"); java.util.regex.Matcher m = p.matcher("hello"); System.out.println(m.replaceFirst("O"));"##,
    );
    assert_eq!(out, vec!["hOllo"]);
}

#[test]
fn matcher_append_replacement_only_first_space() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile(" "); java.util.regex.Matcher m = p.matcher("a b c"); StringBuffer sb = new StringBuffer(); if (m.find()) { m.appendReplacement(sb, "_"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["a_b c"]);
}

#[test]
fn matcher_append_replacement_loop_normalizes_spaces() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile(" +"); java.util.regex.Matcher m = p.matcher("a  b   c"); StringBuffer sb = new StringBuffer(); while (m.find()) { m.appendReplacement(sb, " "); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["a b c"]);
}

#[test]
fn matcher_replace_first_removes_first_nonletter() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("[^a-z]+"); java.util.regex.Matcher m = p.matcher("a--b__c"); System.out.println(m.replaceFirst(""));"##,
    );
    assert_eq!(out, vec!["ab__c"]);
}

#[test]
fn matcher_replace_all_removes_all_nonletter_runs() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("[^a-z]+"); java.util.regex.Matcher m = p.matcher("a--b__c"); System.out.println(m.replaceAll(""));"##,
    );
    assert_eq!(out, vec!["abc"]);
}

#[test]
fn matcher_append_replacement_backreference_swaps_fraction() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("(\\d+)/(\d+)"); java.util.regex.Matcher m = p.matcher("3/4 and 5/6"); StringBuffer sb = new StringBuffer(); if (m.find()) { m.appendReplacement(sb, "$2/$1"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["4/3 and 5/6"]);
}

#[test]
fn matcher_append_replacement_single_match_prefix() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("x"); java.util.regex.Matcher m = p.matcher("xax"); StringBuffer sb = new StringBuffer(); if (m.find()) { m.appendReplacement(sb, "1"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["1ax"]);
}

#[test]
fn matcher_replace_first_on_start_anchor() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("^pre"); java.util.regex.Matcher m = p.matcher("prefix"); System.out.println(m.replaceFirst("post"));"##,
    );
    assert_eq!(out, vec!["postfix"]);
}

#[test]
fn matcher_replace_first_on_word_boundary() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\bjava\b"); java.util.regex.Matcher m = p.matcher("run java now"); System.out.println(m.replaceFirst("vybe"));"##,
    );
    assert_eq!(out, vec!["run vybe now"]);
}

#[test]
fn matcher_append_tail_after_multiple_find_calls() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a"); java.util.regex.Matcher m = p.matcher("aba"); StringBuffer sb = new StringBuffer(); while (m.find()) { m.appendReplacement(sb, "*"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["*b*"]);
}

#[test]
fn matcher_replace_first_parentheses_alternation() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("(cat|dog)"); java.util.regex.Matcher m = p.matcher("dogcat"); System.out.println(m.replaceFirst("pet"));"##,
    );
    assert_eq!(out, vec!["petcat"]);
}

#[test]
fn matcher_append_replacement_empty_deletes_match() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("-"); java.util.regex.Matcher m = p.matcher("a-b"); StringBuffer sb = new StringBuffer(); if (m.find()) { m.appendReplacement(sb, ""); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn matcher_replace_first_dot_matches_first_char() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("."); java.util.regex.Matcher m = p.matcher("abc"); System.out.println(m.replaceFirst("Z"));"##,
    );
    assert_eq!(out, vec!["Zbc"]);
}

#[test]
fn matcher_replace_all_dot_matches_every_char() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("."); java.util.regex.Matcher m = p.matcher("abc"); System.out.println(m.replaceAll("Z"));"##,
    );
    assert_eq!(out, vec!["ZZZ"]);
}

#[test]
fn matcher_append_replacement_builds_parenthesized_words() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\b\\w+\\b"); java.util.regex.Matcher m = p.matcher("go fast"); StringBuffer sb = new StringBuffer(); while (m.find()) { m.appendReplacement(sb, "($0)"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["(go) (fast)"]);
}

#[test]
fn matcher_replace_first_hex_color_channel() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("[0-9a-f]{2}"); java.util.regex.Matcher m = p.matcher("ff00aa"); System.out.println(m.replaceFirst("00"));"##,
    );
    assert_eq!(out, vec!["0000aa"]);
}

#[test]
fn matcher_append_replacement_preserves_trailing_text() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("cat"); java.util.regex.Matcher m = p.matcher("cat!"); StringBuffer sb = new StringBuffer(); if (m.find()) { m.appendReplacement(sb, "dog"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["dog!"]);
}

#[test]
fn matcher_replace_first_optional_u_in_colour() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("colou?r"); java.util.regex.Matcher m = p.matcher("color colour"); System.out.println(m.replaceFirst("hue"));"##,
    );
    assert_eq!(out, vec!["hue colour"]);
}

#[test]
fn matcher_manual_loop_counts_replacements_in_buffer() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d"); java.util.regex.Matcher m = p.matcher("a1b2c3"); StringBuffer sb = new StringBuffer(); int n = 0; while (m.find()) { m.appendReplacement(sb, "X"); n++; } m.appendTail(sb); System.out.println(n); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["3", "aXbXcX"]);
}

#[test]
fn matcher_replace_first_on_alternation() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("red|green|blue"); java.util.regex.Matcher m = p.matcher("green blue red"); System.out.println(m.replaceFirst("yellow"));"##,
    );
    assert_eq!(out, vec!["yellow blue red"]);
}

#[test]
fn matcher_replace_all_on_alternation() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("red|green|blue"); java.util.regex.Matcher m = p.matcher("green blue red"); System.out.println(m.replaceAll("yellow"));"##,
    );
    assert_eq!(out, vec!["yellow yellow yellow"]);
}

#[test]
fn matcher_append_replacement_with_group_zero() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("(ab)+"); java.util.regex.Matcher m = p.matcher("ababx"); StringBuffer sb = new StringBuffer(); if (m.find()) { m.appendReplacement(sb, "($0)"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["(abab)x"]);
}

#[test]
fn matcher_replace_first_case_sensitive_mismatch() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("Java"); java.util.regex.Matcher m = p.matcher("java"); System.out.println(m.replaceFirst("Vybe"));"##,
    );
    assert_eq!(out, vec!["java"]);
}

#[test]
fn matcher_append_replacement_uppercases_each_word() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\w+"); java.util.regex.Matcher m = p.matcher("vybe vm"); StringBuffer sb = new StringBuffer(); while (m.find()) { m.appendReplacement(sb, m.group(0).toUpperCase()); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["VYBE VM"]);
}

#[test]
fn matcher_replace_first_on_plus_quantifier() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a+"); java.util.regex.Matcher m = p.matcher("aaab"); System.out.println(m.replaceFirst("A"));"##,
    );
    assert_eq!(out, vec!["Ab"]);
}

#[test]
fn matcher_replace_all_on_plus_quantifier() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a+"); java.util.regex.Matcher m = p.matcher("aaab"); System.out.println(m.replaceAll("A"));"##,
    );
    assert_eq!(out, vec!["Ab"]);
}

#[test]
fn matcher_append_tail_after_single_replacement() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("b"); java.util.regex.Matcher m = p.matcher("abc"); StringBuffer sb = new StringBuffer(); if (m.find()) { m.appendReplacement(sb, "B"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["aBc"]);
}

#[test]
fn matcher_replace_first_on_escaped_dot() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\."); java.util.regex.Matcher m = p.matcher("a.b.c"); System.out.println(m.replaceFirst(","));"##,
    );
    assert_eq!(out, vec!["a,b.c"]);
}

#[test]
fn matcher_append_replacement_chain_two_passes() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a"); java.util.regex.Matcher m = p.matcher("aba"); StringBuffer sb = new StringBuffer(); while (m.find()) { m.appendReplacement(sb, "o"); } m.appendTail(sb); java.util.regex.Matcher m2 = p.matcher(sb.toString()); StringBuffer sb2 = new StringBuffer(); while (m2.find()) { m2.appendReplacement(sb2, "u"); } m2.appendTail(sb2); System.out.println(sb2.toString());"##,
    );
    assert_eq!(out, vec!["ubu"]);
}

#[test]
fn matcher_append_replacement_skips_when_find_false() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\d+"); java.util.regex.Matcher m = p.matcher("none"); StringBuffer sb = new StringBuffer(); if (m.find()) { m.appendReplacement(sb, "#"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["none"]);
}

#[test]
fn matcher_replace_first_on_star_quantifier() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("x*"); java.util.regex.Matcher m = p.matcher("ab"); System.out.println(m.replaceFirst(""));"##,
    );
    assert_eq!(out, vec!["ab"]);
}

#[test]
fn matcher_append_replacement_uppercases_each_vowel_match() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("[aeiou]"); java.util.regex.Matcher m = p.matcher("hello"); StringBuffer sb = new StringBuffer(); while (m.find()) { m.appendReplacement(sb, m.group().toUpperCase()); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["hEllO"]);
}

#[test]
fn matcher_replace_first_only_touches_initial_whitespace_run() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("\\s+"); java.util.regex.Matcher m = p.matcher("  a  b"); System.out.println(m.replaceFirst("_"));"##,
    );
    assert_eq!(out, vec!["_a  b"]);
}

#[test]
fn matcher_append_replacement_uses_dollar_zero_literal() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("a"); java.util.regex.Matcher m = p.matcher("aba"); StringBuffer sb = new StringBuffer(); while (m.find()) { m.appendReplacement(sb, "\\$0"); } m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["$0b$0"]);
}

#[test]
fn matcher_replace_first_named_group_style_backreference() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("(\\d)(\\d)"); java.util.regex.Matcher m = p.matcher("1234"); System.out.println(m.replaceFirst("$2$1"));"##,
    );
    assert_eq!(out, vec!["2134"]);
}

#[test]
fn matcher_append_tail_preserves_unmatched_prefix() {
    let out = run_main(
        r##"java.util.regex.Pattern p = java.util.regex.Pattern.compile("z"); java.util.regex.Matcher m = p.matcher("prefix"); StringBuffer sb = new StringBuffer(); m.appendTail(sb); System.out.println(sb.toString());"##,
    );
    assert_eq!(out, vec!["prefix"]);
}