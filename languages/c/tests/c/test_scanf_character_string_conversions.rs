//! sscanf character/string conversions — %c %s %[] scansets and width limits.

c_run_cases! {
    sscanf_c_digit_char => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char c; sscanf(\"5\", \"%c\", &c); printf(\"%c\\n\", c); return 0;",
        expect: ["5"]
    },
    sscanf_c_lowercase_z => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char c; sscanf(\"z\", \"%c\", &c); printf(\"%c\\n\", c); return 0;",
        expect: ["z"]
    },
    sscanf_c_does_not_skip_leading_space => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char c; sscanf(\" X\", \"%c\", &c); printf(\"%d\\n\", c); return 0;",
        expect: ["32"]
    },
    sscanf_c_second_char_after_space => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char a,b; sscanf(\" X\", \"%c%c\", &a, &b); printf(\"%c\\n\", b); return 0;",
        expect: ["X"]
    },
    sscanf_c_tab_character => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char c; sscanf(\"\\t\", \"%c\", &c); printf(\"%d\\n\", c); return 0;",
        expect: ["9"]
    },
    sscanf_c_newline_character => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char c; sscanf(\"\\n\", \"%c\", &c); printf(\"%d\\n\", c); return 0;",
        expect: ["10"]
    },
    sscanf_c_two_chars_sequential => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char a,b; sscanf(\"pq\", \"%c%c\", &a, &b); printf(\"%c%c\\n\", a, b); return 0;",
        expect: ["pq"]
    },
    sscanf_c_after_parsed_integer => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; char c; sscanf(\"9z\", \"%d%c\", &n, &c); printf(\"%d %c\\n\", n, c); return 0;",
        expect: ["9 z"]
    },
    sscanf_c_width_reads_two => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[4]; sscanf(\"wxy\", \"%2c\", buf); printf(\"%c%c\\n\", buf[0], buf[1]); return 0;",
        expect: ["wx"]
    },
    sscanf_c_skip_first_via_literal => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char c; sscanf(\"Qm\", \"Q%c\", &c); printf(\"%c\\n\", c); return 0;",
        expect: ["m"]
    },
    sscanf_c_punctuation_exclamation => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char c; sscanf(\"!\", \"%c\", &c); printf(\"%c\\n\", c); return 0;",
        expect: ["!"]
    },
    sscanf_c_uppercase_m => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char c; sscanf(\"M\", \"%c\", &c); printf(\"%c\\n\", c); return 0;",
        expect: ["M"]
    },
    sscanf_c_three_char_run => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char a,b,c; sscanf(\"abc\", \"%c%c%c\", &a, &b, &c); printf(\"%c%c%c\\n\", a, b, c); return 0;",
        expect: ["abc"]
    },
    sscanf_c_after_literal_hash => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char c; sscanf(\"#Z\", \"#%c\", &c); printf(\"%c\\n\", c); return 0;",
        expect: ["Z"]
    },
    sscanf_c_percent_sign => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char c; sscanf(\"%\", \"%c\", &c); printf(\"%c\\n\", c); return 0;",
        expect: ["%"]
    },
    sscanf_s_basic_token => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"vybe\", \"%s\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["vybe"]
    },
    sscanf_s_skips_leading_whitespace => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"  token\", \"%s\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["token"]
    },
    sscanf_s_width_five_on_longer => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"compiler\", \"%5s\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["compi"]
    },
    sscanf_s_width_two_alpha => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[8]; sscanf(\"alpha\", \"%2s\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["al"]
    },
    sscanf_s_two_tokens => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char a[8], b[8]; sscanf(\"foo bar\", \"%s %s\", a, b); printf(\"%s-%s\\n\", a, b); return 0;",
        expect: ["foo-bar"]
    },
    sscanf_s_stops_at_whitespace => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"one two\", \"%s\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["one"]
    },
    sscanf_s_underscore_token => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"snake_case\", \"%s\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["snake_case"]
    },
    sscanf_s_digits_as_string => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[8]; sscanf(\"4096\", \"%s\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["4096"]
    },
    sscanf_s_after_integer => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "int n; char buf[8]; sscanf(\"12items\", \"%d%s\", &n, buf); printf(\"%d %s\\n\", n, buf); return 0;",
        expect: ["12 items"]
    },
    sscanf_s_width_three_on_vybe => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[8]; sscanf(\"vybe-lang\", \"%3s\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["vyb"]
    },
    sscanf_s_mixed_with_char => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[8]; char c; sscanf(\"ok!\", \"%s%c\", buf, &c); printf(\"%s %c\\n\", buf, c); return 0;",
        expect: ["ok !"]
    },
    sscanf_s_second_token_after_space => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char a[8],b[8]; sscanf(\"skip rest\", \"%s %s\", a,b); printf(\"%s\\n\", b); return 0;",
        expect: ["rest"]
    },
    sscanf_s_single_char_token => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[4]; sscanf(\"x\", \"%s\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["x"]
    },
    sscanf_s_hyphenated_word => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"well-formed\", \"%s\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["well-formed"]
    },
    sscanf_s_three_tokens => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char a[6],b[6],c[6]; sscanf(\"red green blue\", \"%s %s %s\", a,b,c); printf(\"%s/%s/%s\\n\", a,b,c); return 0;",
        expect: ["red/green/blue"]
    },
    sscanf_scanset_digits_only => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"123abc\", \"%[0-9]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["123"]
    },
    sscanf_scanset_lowercase_letters => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"vybe123\", \"%[a-z]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["vybe"]
    },
    sscanf_scanset_negated_comma => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"field,rest\", \"%[^,]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["field"]
    },
    sscanf_scanset_vowels_only => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[8]; sscanf(\"aeiouxyz\", \"%[aeiou]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["aeiou"]
    },
    sscanf_scanset_hex_digits => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"deadbeefg\", \"%[0-9a-f]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["deadbeef"]
    },
    sscanf_scanset_negated_digits => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"abc123\", \"%[^0-9]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["abc"]
    },
    sscanf_scanset_width_four_digits => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[8]; sscanf(\"987654\", \"%4[0-9]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["9876"]
    },
    sscanf_scanset_upper_hex => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[8]; sscanf(\"ABCDxy\", \"%[A-F]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["ABCD"]
    },
    sscanf_scanset_until_space => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"word tail\", \"%[^ ]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["word"]
    },
    sscanf_scanset_specific_set_abc => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[8]; sscanf(\"abx\", \"%[abc]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["ab"]
    },
    sscanf_scanset_negated_space => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"nospace here\", \"%[^ ]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["nospace"]
    },
    sscanf_scanset_line_without_newline => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"line\\n\", \"%[^\\n]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["line"]
    },
    sscanf_scanset_punctuation_run => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"!!!text\", \"%[!]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["!!!"]
    },
    sscanf_scanset_mixed_alpha_num => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"a1b2c\", \"%[a-z0-9]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["a1b2c"]
    },
    sscanf_scanset_after_literal_dash => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[8]; sscanf(\"-42\", \"-%[0-9]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["42"]
    },
    sscanf_scanset_negated_equals => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[16]; sscanf(\"key=value\", \"%[^=]\", buf); printf(\"%s\\n\", buf); return 0;",
        expect: ["key"]
    },
    sscanf_scanset_pair_fields => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char a[8],b[8]; sscanf(\"aa,bb\", \"%[a-z],%[a-z]\", a,b); printf(\"%s %s\\n\", a,b); return 0;",
        expect: ["aa bb"]
    },
    sscanf_scanset_empty_when_mismatch => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char buf[8]; int n=sscanf(\"9\", \"%[a-z]\", buf); buf[0]=0; printf(\"%d\\n\", n); return 0;",
        expect: ["0"]
    },
    sscanf_s_and_scanset_combo => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char a[8],b[8]; sscanf(\"12:xy\", \"%[0-9]:%[a-z]\", a,b); printf(\"%s %s\\n\", a,b); return 0;",
        expect: ["12 xy"]
    },
    sscanf_c_and_s_combo => {
        includes: ["<stdio.h>"],
        decls: "",
        body: "char c; char buf[8]; sscanf(\"Zdata\", \"%c%s\", &c, buf); printf(\"%c %s\\n\", c, buf); return 0;",
        expect: ["Z data"]
    },
}
