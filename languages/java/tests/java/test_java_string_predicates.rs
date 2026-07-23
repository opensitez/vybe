use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(is_empty_true, "System.out.println(\"\".isEmpty());", "true");
jt!(
    is_empty_false,
    "System.out.println(\"a\".isEmpty());",
    "false"
);
jt!(length_zero, "System.out.println(\"\".length());", "0");
jt!(length_three, "System.out.println(\"abc\".length());", "3");
jt!(
    starts_with_true,
    "System.out.println(\"apple\".startsWith(\"app\"));",
    "true"
);
jt!(
    starts_with_false,
    "System.out.println(\"apple\".startsWith(\"appx\"));",
    "false"
);
jt!(
    ends_with_true,
    "System.out.println(\"apple\".endsWith(\"ple\"));",
    "true"
);
jt!(
    ends_with_false,
    "System.out.println(\"apple\".endsWith(\"apl\"));",
    "false"
);
jt!(
    contains_true,
    "System.out.println(\"banana\".contains(\"ana\"));",
    "true"
);
jt!(
    contains_false,
    "System.out.println(\"banana\".contains(\"abc\"));",
    "false"
);
jt!(
    index_of_first,
    "System.out.println(\"banana\".indexOf(\"na\"));",
    "2"
);
jt!(
    index_of_missing,
    "System.out.println(\"banana\".indexOf(\"zz\"));",
    "-1"
);
jt!(
    last_index_of,
    "System.out.println(\"banana\".lastIndexOf(\"na\"));",
    "4"
);
jt!(char_at_zero, "System.out.println(\"abc\".charAt(0));", "a");
jt!(char_at_last, "System.out.println(\"abc\".charAt(2));", "c");
jt!(
    equals_true,
    "System.out.println(\"abc\".equals(\"abc\"));",
    "true"
);
jt!(
    equals_false,
    "System.out.println(\"abc\".equals(\"ABC\"));",
    "false"
);
jt!(
    equals_ignore_case_true,
    "System.out.println(\"abc\".equalsIgnoreCase(\"ABC\"));",
    "true"
);
jt!(
    compare_to_lower,
    "System.out.println(\"abc\".compareTo(\"abd\"));",
    "-1"
);
jt!(
    compare_to_equal,
    "System.out.println(\"abc\".compareTo(\"abc\"));",
    "0"
);
jt!(
    compare_to_higher,
    "System.out.println(\"abc\".compareTo(\"abb\"));",
    "1"
);
jt!(
    region_matches_true,
    "System.out.println(\"abcdef\".regionMatches(2, \"cd\", 0, 2));",
    "true"
);
jt!(
    region_matches_false,
    "System.out.println(\"abcdef\".regionMatches(2, \"ef\", 0, 2));",
    "false"
);
jt!(
    matches_simple,
    "System.out.println(\"123\".matches(\"[0-9]+\"));",
    "true"
);
jt!(
    replace_all_not,
    "System.out.println(\"a1b2\".replace('a', 'z'));",
    "z1b2"
);
jt!(
    substring_front,
    "System.out.println(\"hello\".substring(2));",
    "llo"
);
jt!(
    substring_mid,
    "System.out.println(\"hello\".substring(1, 4));",
    "ell"
);
jt!(
    to_char_array_len,
    "System.out.println(\"cat\".toCharArray().length);",
    "3"
);
jt!(
    hashcode_stable,
    "System.out.println(\"a\".hashCode() == \"a\".hashCode());",
    "true"
);
