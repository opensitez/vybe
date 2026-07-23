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
    to_upper,
    "System.out.println(\"abc\".toUpperCase());",
    "ABC"
);
jt!(
    to_lower,
    "System.out.println(\"ABC\".toLowerCase());",
    "abc"
);
jt!(trim_start, "System.out.println(\"  a\".trim());", "a");
jt!(trim_both, "System.out.println(\"  abc  \".trim());", "abc");
jt!(
    replace_text,
    "System.out.println(\"a-b-c\".replace('-', '_'));",
    "a_b_c"
);
jt!(
    replace_char,
    "System.out.println(\"abc\".replace('a', 'z'));",
    "zbc"
);
jt!(
    substring_and_concat,
    "String s = \"abcdef\"; System.out.println(s.substring(1, 4));",
    "bcd"
);
jt!(
    value_of_int,
    "System.out.println(String.valueOf(12));",
    "12"
);
jt!(
    value_of_bool,
    "System.out.println(String.valueOf(true));",
    "true"
);
jt!(
    value_of_char,
    "System.out.println(String.valueOf('x'));",
    "x"
);
jt!(
    format_empty,
    "System.out.println(String.format(\"%d\", 7));",
    "7"
);
jt!(
    replace_first,
    "System.out.println(\"foofoo\".replaceFirst(\"foo\", \"bar\"));",
    "barfoo"
);
jt!(
    starts_with_prefix,
    "System.out.println(\"hello\".startsWith(\"he\"));",
    "true"
);
jt!(
    ends_with_suffix,
    "System.out.println(\"hello\".endsWith(\"lo\"));",
    "true"
);
jt!(
    concat_plus,
    "System.out.println(\"ab\".concat(\"cd\"));",
    "abcd"
);
jt!(
    compare_ignore_case,
    "System.out.println(\"HeLLo\".compareToIgnoreCase(\"hello\"));",
    "0"
);
jt!(
    repeat_three,
    "System.out.println(\"ab\".repeat(2));",
    "abab"
);
jt!(
    replace_all_like,
    "System.out.println(\"banana\".replace(\"na\", \"NA\"));",
    "baNANA"
);
jt!(
    split_count,
    "String[] p = \"a,b,c\".split(\",\"); System.out.println(p.length);",
    "3"
);
jt!(
    split_mid,
    "String[] p = \"a,b,c\".split(\",\"); System.out.println(p[1]);",
    "b"
);
jt!(
    join_two,
    "System.out.println(String.join(\"-\", \"a\", \"b\", \"c\"));",
    "a-b-c"
);
jt!(
    chars_length,
    "System.out.println(\"abc\".chars().count());",
    "3"
);
jt!(
    to_char_array_last,
    "char[] a = \"xyz\".toCharArray(); System.out.println(a[a.length - 1]);",
    "z"
);
jt!(
    code_point_at,
    "System.out.println(\"a\".codePointAt(0));",
    "97"
);
jt!(
    interned,
    "String a = new String(\"x\").intern(); System.out.println(a.intern().equals(\"x\"));",
    "true"
);
jt!(
    strip_not_blank,
    "System.out.println(\" x \".strip().length());",
    "1"
);
jt!(
    format_two,
    "System.out.println(String.format(\"%d:%s\", 2, \"x\"));",
    "2:x"
);
jt!(
    index_of_char,
    "System.out.println(\"abcd\".indexOf('c'));",
    "2"
);
jt!(
    replace_literal,
    "System.out.println(\"hello world\".replace(\"world\", \"java\"));",
    "hello java"
);
jt!(
    compare_region,
    "System.out.println(\"abcdef\".regionMatches(false, 2, \"CD\", 0, 2));",
    "false"
);
