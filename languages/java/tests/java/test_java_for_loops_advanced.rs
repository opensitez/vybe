use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(for_sum_0_to_4, "int s = 0; for (int i = 0; i < 5; i++) { s += i; } System.out.println(s);", "10");
jt!(for_skip_even, "int s = 0; for (int i = 0; i < 7; i++) { if ((i & 1) == 1) s += i; } System.out.println(s);", "9");
jt!(for_with_continue, "int s = 0; for (int i = 0; i < 6; i++) { if (i == 3) continue; s += i; } System.out.println(s);", "12");
jt!(for_with_break, "int s = 0; for (int i = 0; i < 10; i++) { s += i; if (i == 4) break; } System.out.println(s);", "10");
jt!(for_multi_init, "int s = 0; for (int i = 0, j = 2; i < 3; i++, j += 2) { s += j; } System.out.println(s);", "12");
jt!(for_descending, "int s = 0; for (int i = 5; i > 0; i--) { s += i; } System.out.println(s);", "15");
jt!(for_step_two, "int s = 0; for (int i = 0; i < 10; i += 2) { s += i; } System.out.println(s);", "20");
jt!(for_step_negative, "int s = 0; for (int i = 10; i > 5; i -= 2) { s += i; } System.out.println(s);", "24");
jt!(for_nested_outer, "int s = 0; for (int i = 0; i < 3; i++) { for (int j = 0; j < 2; j++) { s += (i * 2) + j; } } System.out.println(s);", "9");
jt!(for_nested_break, "int s = 0; for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { if (j == 1) break; s += i + j; } } System.out.println(s);", "3");
jt!(for_with_labels, "int s = 0; for (int i = 0; i < 3; i++) { s += i; } for (int i = 3; i < 6; i++) { s += i; } System.out.println(s);", "15");
jt!(for_boolean_expression, "int s = 0; boolean flag = true; for (int i = 0; flag; i++) { s += i; if (i == 2) flag = false; } System.out.println(s);", "3");
jt!(for_in_array, "int[] v = {1, 2, 3}; int s = 0; for (int i = 0; i < v.length; i++) { s += v[i]; } System.out.println(s);", "6");
jt!(for_if_inside, "int s = 0; for (int i = 0; i < 6; i++) { if (i == 4) continue; s += i; } System.out.println(s);", "11");
jt!(for_post_expression_effect, "int s = 0; for (int i = 0; i < 4; s += i, i++) {} System.out.println(s);", "6");
jt!(for_char_range, "int s = 0; for (char c = 'a'; c <= 'c'; c++) { s += c; } System.out.println(s - 291);", "3");
jt!(for_string_concat, "String s = \"\"; for (int i = 0; i < 3; i++) { s += i; } System.out.println(s);", "012");
jt!(for_complex_update, "int s = 1; for (int i = 1; i < 4; i *= 2) { s *= i; } System.out.println(s);", "3");
jt!(for_without_body_empty, "int c = 0; for (int i = 0; i < 4; i++) c++; System.out.println(c);", "4");
jt!(for_outer_multiplication, "int s = 1; for (int i = 1; i <= 4; i++) { s *= i; } System.out.println(s);", "24");
jt!(for_condition_changes, "int c = 0; for (int i = 0; i < 10; i++) { c += i; if (c > 10) break; } System.out.println(c);", "10");
jt!(for_with_boolean_cast, "int c = 0; for (int i = 0; i < 4; i++) { if ((i & 1) == 0) c++; } System.out.println(c);", "2");
jt!(for_multiple_statements, "int c = 0; for (int i = 0, j = 0; i < 3; i++, j++) { c += j; } System.out.println(c);", "3");
jt!(for_reverse_index, "int c = 0; for (int i = 4; i >= 0; i--) { c += i; } System.out.println(c);", "10");
jt!(for_with_local_array, "int[] a = {2, 4, 6}; int c = 0; for (int i = 0; i < a.length; i++) { c += a[i] / 2; } System.out.println(c);", "6");
jt!(for_string_len, "String s = \"abc\"; int c = 0; for (int i = 0; i < s.length(); i++) { c += s.charAt(i); } System.out.println(c - 291);", "3");
jt!(for_nested_while_equiv, "int c = 0; int i = 0; for (; i < 3; i++) { c += i; } System.out.println(c);", "3");
jt!(for_ternary_limit, "int c = 0; for (int i = 0; i < (2 > 1 ? 4 : 1); i++) { c += i; } System.out.println(c);", "6");
jt!(for_zero_times, "int c = 0; for (int i = 0; i < 0; i++) { c++; } System.out.println(c);", "0");

