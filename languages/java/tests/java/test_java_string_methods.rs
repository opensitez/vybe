use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(string_length_simple, "System.out.println(\"abc\".length());", "3");
jt!(string_is_empty_false, "System.out.println(\"abc\".isEmpty());", "false");
jt!(string_is_empty_true, "System.out.println(\"\".isEmpty());", "true");
jt!(string_char_at_zero, "System.out.println(\"cat\".charAt(0));", "c");
jt!(string_char_at_last, "System.out.println(\"cat\".charAt(2));", "t");
jt!(string_substring_start, "System.out.println(\"hello\".substring(2));", "llo");
jt!(string_substring_range, "System.out.println(\"hello\".substring(1,4));", "ell");
jt!(string_to_upper, "System.out.println(\"java\".toUpperCase());", "JAVA");
jt!(string_to_lower, "System.out.println(\"JaVa\".toLowerCase());", "java");
jt!(string_trim_spaces, "System.out.println(\"  hi  \".trim());", "hi");
jt!(string_concat_plus, "System.out.println(\"a\" + \"b\" + 3);", "ab3");
jt!(string_contains_true, "System.out.println(\"dynamic\".contains(\"nam\"));", "true");
jt!(string_contains_false, "System.out.println(\"dynamic\".contains(\"nom\"));", "false");
jt!(string_starts_with, "System.out.println(\"java\".startsWith(\"ja\"));", "true");
jt!(string_ends_with, "System.out.println(\"java\".endsWith(\"va\"));", "true");
jt!(string_replace_char, "System.out.println(\"cava\".replace('a','o'));", "covo");
jt!(string_index_of_char, "System.out.println(\"banana\".indexOf('n'));", "2");
jt!(string_last_index_of, "System.out.println(\"banana\".lastIndexOf('a'));", "5");
jt!(string_index_of_substring, "System.out.println(\"abracadabra\".indexOf(\"cad\"));", "3");
jt!(string_compare_to_true, "System.out.println(\"aa\".compareTo(\"ab\"));", "-1");
jt!(string_compare_to_false, "System.out.println(\"ab\".compareTo(\"aa\"));", "1");
jt!(string_compare_to_equal, "System.out.println(\"same\".compareTo(\"same\"));", "0");
jt!(string_split_parts, "System.out.println(java.util.Arrays.toString(\"a,b,c\".split(\"\\,\"))[1]);", "b");
jt!(string_format_integer, "System.out.println(String.format(\"%d\", 42));", "42");
jt!(string_region_matches_true, "System.out.println(\"abcdef\".regionMatches(1, \"bc\", 0, 2));", "true");
jt!(string_region_matches_false, "System.out.println(\"abcdef\".regionMatches(1, \"bd\", 0, 2));", "false");
jt!(string_equals_ignore_case_true, "System.out.println(\"Java\".equalsIgnoreCase(\"jAvA\"));", "true");
jt!(string_repeat_twice, "System.out.println(\"x\".repeat(3));", "xxx");
jt!(string_concat_method, "System.out.println(\"x\".concat(\"y\"));", "xy");
