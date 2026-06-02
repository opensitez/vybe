use super::helpers::{compile_ok, run_ruby_one};

// -- Regexp.new constructor

#[test]
fn regexp_new_constructor() {
    compile_ok(
        r#"r = Regexp.new('hello')
"#,
    );
}

// -- Regexp literal /pattern/

#[test]
fn regexp_literal_pattern() {
    compile_ok(
        "r = /hello/
",
    );
}

// -- Regexp with i flag (case-insensitive)

#[test]
fn regexp_flag_case_insensitive() {
    compile_ok(
        r#"r = /hello/i
x = 'HELLO' =~ r
"#,
    );
}

// -- Regexp with m flag (dot matches newline)

#[test]
fn regexp_flag_multiline_dot() {
    compile_ok(
        "r = /a.b/m
",
    );
}

// -- Regexp with x flag (extended, whitespace ignored)

#[test]
fn regexp_flag_extended() {
    compile_ok(
        "r = /hello   # match hello/x
",
    );
}

// -- =~ operator stores result in dollar-tilde

#[test]
fn regexp_match_op_stores_match_data() {
    compile_ok(
        r#"'hello' =~ /ell/
m = $~
"#,
    );
}

// -- Capture group dollar variables $1, $2

#[test]
fn regexp_capture_group_dollar_vars() {
    compile_ok(
        r#"'2024-01-15' =~ /(\d{4})-(\d{2})/
y = $1
m = $2
"#,
    );
}

// -- $& matched string special variable

#[test]
fn regexp_dollar_ampersand_matched_string() {
    compile_ok(
        r#"'hello world' =~ /w\w+/
matched = $&
"#,
    );
}

// -- Named capture group (?<name>...)

#[test]
fn regexp_named_capture_group() {
    compile_ok(
        r#"m = 'John 30'.match(/(?<name>\w+) (?<age>\d+)/)
"#,
    );
}

// -- match returns MatchData object

#[test]
fn regexp_match_returns_matchdata() {
    compile_ok(
        r#"md = 'hello'.match(/e(l+)/)
"#,
    );
}

// -- MatchData index access

#[test]
fn matchdata_index_access() {
    compile_ok(
        r#"md = 'hello'.match(/(e)(l+)/)
full = md[0]
first = md[1]
"#,
    );
}

// -- MatchData named access

#[test]
fn matchdata_named_access() {
    compile_ok(
        r#"md = 'John'.match(/(?<first>\w+)/)
n = md[:first]
"#,
    );
}

// -- MatchData pre_match and post_match

#[test]
fn matchdata_pre_and_post_match() {
    compile_ok(
        r#"md = 'hello world'.match(/wor/)
pre = md.pre_match
post = md.post_match
"#,
    );
}

// -- match? returns bool, no side effects

#[test]
fn regexp_match_predicate_no_side_effects() {
    compile_ok(
        r#"result = 'hello'.match?(/ell/)
"#,
    );
}

// -- gsub with regex and block

#[test]
fn regexp_gsub_with_block() {
    compile_ok(
        r#"result = 'hello world'.gsub(/\w+/) { |w| w.upcase }
"#,
    );
}

// -- gsub with hash replacement

#[test]
fn regexp_gsub_with_hash() {
    compile_ok(
        r#"result = 'aeiou'.gsub(/[aeiou]/, 'a' => '1', 'e' => '2', 'i' => '3')
"#,
    );
}

// -- scan with capture groups returns nested array

#[test]
fn regexp_scan_capture_groups_nested() {
    compile_ok(
        r#"result = 'one 1, two 2'.scan(/(\w+) (\d)/)
"#,
    );
}

// -- split with capturing regex

#[test]
fn regexp_split_with_capturing_regex() {
    compile_ok(
        r#"result = 'a1b2c3'.split(/(\d)/)
"#,
    );
}

// -- sub with backreference in replacement

#[test]
fn regexp_sub_backreference_in_replacement() {
    compile_ok(
        r#"result = 'hello world'.sub(/(\w+)/, '[\1]')
"#,
    );
}

// -- =~ returns nil on no match

#[test]
fn regexp_match_op_nil_on_no_match() {
    compile_ok(
        r#"result = 'hello' =~ /xyz/
"#,
    );
}

#[test]
fn regexp_match_op_nil_runtime() {
    assert_eq!(
        run_ruby_one(
            r#"puts ('hello' =~ /xyz/).nil?
"#
        ),
        "true"
    );
}

// -- Regexp#match vs String#match

#[test]
fn regexp_match_vs_string_match() {
    compile_ok(
        r#"r = /\d+/
md1 = r.match('abc 42 def')
md2 = 'abc 42 def'.match(r)
"#,
    );
}

// -- Regexp#=== for case/when

#[test]
fn regexp_triple_eq_case_when() {
    compile_ok(
        r#"s = 'hello123'
result = case s
when /^\d/ then 'digit'
when /^[a-z]/ then 'letter'
else 'other'
end
"#,
    );
}

// -- Character class [a-z]

#[test]
fn regexp_character_class() {
    compile_ok(
        r#"result = 'Hello World'.gsub(/[a-z]/, '*')
"#,
    );
}

// -- Anchors ^, $, \A, \Z

#[test]
fn regexp_anchors() {
    compile_ok(
        r#"a = 'hello' =~ /\Ahello\Z/
b = 'hello' =~ /^hello$/
"#,
    );
}

// -- Regexp#source returns pattern string

#[test]
fn regexp_source_returns_pattern() {
    compile_ok(
        r#"src = /hello\d+/i.source
"#,
    );
}

#[test]
fn regexp_source_runtime() {
    assert_eq!(
        run_ruby_one(
            r#"puts /hello/.source
"#
        ),
        "hello"
    );
}
