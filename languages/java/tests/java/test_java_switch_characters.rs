use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(
    switch_char_a,
    "char c = 'a'; int v = 0; switch (c) { case 'a': v = 1; break; case 'b': v = 2; break; default: v = 0; } System.out.println(v);",
    "1"
);
jt!(
    switch_char_b,
    "char c = 'b'; int v = 0; switch (c) { case 'a': v = 1; break; case 'b': v = 2; break; default: v = 0; } System.out.println(v);",
    "2"
);
jt!(
    switch_char_default,
    "char c = 'z'; int v = 0; switch (c) { case 'a': v = 1; break; case 'b': v = 2; break; default: v = 3; } System.out.println(v);",
    "3"
);
jt!(
    switch_char_numeric_addition,
    "char c = 'c'; int s = 0; switch (c) { case 'a': s = 1; break; case 'c': s = c; break; default: s = 0; } System.out.println(s - 96);",
    "3"
);
jt!(
    switch_char_word,
    "char c = 'x'; String r = \"\"; switch (c) { case 'w': r = \"w\"; break; case 'x': r = \"x\"; break; case 'y': r = \"y\"; break; default: r = \"z\"; } System.out.println(r);",
    "x"
);
jt!(
    switch_vowel_aeio,
    "char c = 'e'; String r = \"\"; switch (c) { case 'a': case 'e': case 'i': case 'o': case 'u': r = \"vowel\"; break; default: r = \"consonant\"; } System.out.println(r);",
    "vowel"
);
jt!(
    switch_digit_char,
    "char c = '7'; int v = 0; switch (c) { case '0': case '1': case '2': v = 1; break; case '7': v = 7; break; default: v = -1; } System.out.println(v);",
    "7"
);
jt!(
    switch_unicode_char,
    r#"char c = '\u0041'; int v = 0; switch (c) { case 'A': v = 1; break; case 'B': v = 2; break; default: v = 0; } System.out.println(v);"#,
    "1"
);
jt!(
    switch_symbol,
    "char c = '!'; int v = 0; switch (c) { case '!': v = 1; break; case '?': v = 2; break; default: v = 0; } System.out.println(v);",
    "1"
);
jt!(
    switch_newline_escape,
    "char c = '\\n'; int v = 0; switch (c) { case '\\n': v = 1; break; case '\\t': v = 2; break; default: v = 0; } System.out.println(v);",
    "1"
);
jt!(
    switch_space,
    "char c = ' '; String s = \"space\"; switch (c) { case ' ': s = \"space\"; break; case '\\t': s = \"tab\"; break; default: s = \"other\"; } System.out.println(s);",
    "space"
);
jt!(
    switch_upper_lower,
    "char c = 'F'; String s = \"\"; switch (c) { case 'F': s = \"up\"; break; case 'f': s = \"low\"; break; default: s = \"none\"; } System.out.println(s);",
    "up"
);
jt!(
    switch_pairing,
    "char c = 'h'; int a = 0; switch (c) { case 'g': a = 1; break; case 'h': a = 2; break; case 'i': a = 3; break; default: a = 4; } System.out.println(a);",
    "2"
);
jt!(
    switch_many,
    "char c = 'm'; int a = 0; switch (c) { case 'k': a = 1; case 'l': a = 2; case 'm': a = 3; default: a = 4; } System.out.println(a);",
    "4"
);
jt!(
    switch_in_loop_chars,
    "char[] c = {'a','b','c'}; int s = 0; for (int i = 0; i < c.length; i++) { switch (c[i]) { case 'a': s += 1; break; case 'b': s += 2; break; default: s += 3; } } System.out.println(s);",
    "6"
);
jt!(
    switch_with_escape,
    "char c = '\\t'; int v = 0; switch (c) { case '\\n': v = 1; break; case '\\t': v = 2; break; default: v = 0; } System.out.println(v);",
    "2"
);
jt!(
    switch_derived_index,
    "char c = 'd'; int idx = c - 'a'; int v = 0; switch (idx) { case 3: v = 3; break; default: v = 0; } System.out.println(v);",
    "3"
);
jt!(
    switch_char_and_string,
    "char c = 'b'; String s = \"\"; switch (c) { case 'a': s = \"1\"; break; case 'b': s = \"2\"; break; default: s = \"3\"; } System.out.println(s);",
    "2"
);
jt!(
    switch_unicode_hex,
    "char c = '\\u0042'; int v = 0; switch (c) { case 'B': v = 2; break; default: v = 0; } System.out.println(v);",
    "2"
);
jt!(
    switch_math_expr,
    "char c = 'c'; int x = 0; switch (c) { case 'c': x = 100 / 10; break; default: x = 0; } System.out.println(x);",
    "10"
);
jt!(
    switch_with_empty_case,
    "char c = 'q'; int v = 0; switch (c) { case 'p': v = 1; break; case 'q': case 'r': v = 2; break; default: v = 0; } System.out.println(v);",
    "2"
);
jt!(
    switch_char_chain_no_break,
    "char c = 'a'; int v = 0; switch (c) { case 'a': v += 1; case 'b': v += 2; case 'c': v += 3; default: v += 4; } System.out.println(v);",
    "10"
);
jt!(
    switch_char_nested,
    "char c = 'e'; int v = 0; switch (c) { case 'a': v = 1; break; case 'e': switch (c) { case 'e': v = 5; break; default: v = 0; } break; default: v = 3; } System.out.println(v);",
    "5"
);
jt!(
    switch_empty_default_only,
    "char c = 'z'; int v = 0; switch (c) { default: v = 9; } System.out.println(v);",
    "9"
);
jt!(
    switch_large_ascii,
    "char c = '\\u00FF'; int v = 0; switch (c) { case '\\u00FF': v = 1; break; default: v = 0; } System.out.println(v);",
    "1"
);
jt!(
    switch_on_char_array,
    "char[] c = {'x','y'}; int total = 0; for (int i = 0; i < c.length; i++) { switch (c[i]) { case 'x': total += 1; break; case 'y': total += 2; break; default: total += 3; } } System.out.println(total);",
    "3"
);
jt!(
    switch_char_ternary_mix,
    "char c = 'a'; String s = \"\"; switch (c) { case 'a': s = \"a\"; break; case 'b': s = \"b\"; break; default: s = \"z\"; } System.out.println(s.equals(\"a\") ? 1 : 0);",
    "1"
);
