use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(sum_ints, "int[] values = {1,2,3}; int s = 0; for (int v : values) { s += v; } System.out.println(s);", "6");
jt!(sum_from_expression, "int[] values = {1,2,3,4}; int s = 0; for (int v : values) { s += v * 2; } System.out.println(s);", "20");
jt!(count_items, "int[] values = {1,2,3,4}; int s = 0; for (int v : values) { s++; } System.out.println(s);", "4");
jt!(sum_if_even, "int[] values = {1,2,3,4,5}; int s = 0; for (int v : values) { if (v % 2 == 0) s += v; } System.out.println(s);", "6");
jt!(sum_if_odd, "int[] values = {1,2,3,4,5}; int s = 0; for (int v : values) { if (v % 2 != 0) s += v; } System.out.println(s);", "9");
jt!(product, "int[] values = {1,2,3,4}; int p = 1; for (int v : values) { p *= v; } System.out.println(p);", "24");
jt!(min_value, "int[] values = {5,2,8,1}; int m = values[0]; for (int v : values) { if (v < m) m = v; } System.out.println(m);", "1");
jt!(max_value, "int[] values = {5,2,8,1}; int m = values[0]; for (int v : values) { if (v > m) m = v; } System.out.println(m);", "8");
jt!(concat_chars, "char[] chars = {'a','b','c'}; String s = \"\"; for (char c : chars) s += c; System.out.println(s);", "abc");
jt!(empty_loop_sum, "int[] values = {}; int s = 0; for (int v : values) { s += v; } System.out.println(s);", "0");
jt!(nested_foreach_sum, "int[][] matrix = {{1,2},{3,4}}; int s=0; for (int[] row : matrix) { for (int v : row) { s += v; } } System.out.println(s);", "10");
jt!(nested_string_lengths, "String[] words = {\"a\", \"bc\", \"def\"}; int s = 0; for (String w : words) { s += w.length(); } System.out.println(s);", "6");
jt!(indexing_like_counter, "int[] values = {4,5,6}; int i=0; int s=0; for (int v : values) { s += v + i; i++; } System.out.println(s);", "18");
jt!(foreach_continue_like, "int[] values = {1,2,3,4,5}; int s = 0; for (int v : values) { if (v == 3) continue; s += v; } System.out.println(s);", "12");
jt!(foreach_breakish_flag, "int[] values = {1,2,3,4}; int s = 0; for (int v : values) { if (v == 3) break; s += v; } System.out.println(s);", "3");
jt!(sum_as_int, "int[] values = {1,2,3}; int s = 0; for (int v : values) s = s + v; System.out.println(s);", "6");
jt!(average_division, "int[] values = {2,4,6}; int s = 0; int c = 0; for (int v : values) { s += v; c++; } System.out.println(s / c);", "4");
jt!(all_positive, "int[] values = {1,2,3}; boolean ok = true; for (int v : values) { if (v <= 0) ok = false; } System.out.println(ok);", "true");
jt!(contains_three, "int[] values = {1,2,3,4}; boolean found = false; for (int v : values) { if (v == 3) found = true; } System.out.println(found);", "true");
jt!(sum_strings, "String[] words = {\"a\", \"b\", \"c\"}; String s = \"\"; for (String w : words) s += w; System.out.println(s);", "abc");
jt!(count_with_continue, "int[] values = {1,2,3,4,5}; int c = 0; for (int v : values) { if (v % 2 == 0) continue; c++; } System.out.println(c);", "3");
jt!(reduce_like_add, "int[] values = {5,1,1}; int s = 1; for (int v : values) { s *= v; } System.out.println(s);", "5");
jt!(repeat_three, "int[] values = {1,1,1}; int s = 0; for (int v : values) { s += 3 * v; } System.out.println(s);", "9");
jt!(array_length_from_loop, "int[] values = {8,9}; int c=0; for (int v : values) c++; System.out.println(c == values.length ? c : 0);", "2");
jt!(sum_abs, "int[] values = {-1,2,-3,4}; int s=0; for (int v : values) { s += v < 0 ? -v : v; } System.out.println(s);", "10");
jt!(join_nontrivial, "String[] words = {\"x\", \"y\", \"z\"}; String s = words[0]; for (String w : words) { if (w != words[0]) s += \"|\" + w; } System.out.println(s);", "x|y|z");
jt!(sum_until_condition, "int[] values = {1,2,3,4}; int s = 0; for (int v : values) { s += v; if (s > 4) break; } System.out.println(s);", "6");
jt!(sum_char_codes, "char[] values = {'A', 'B'}; int s = 0; for (char c : values) { s += c; } System.out.println(s);", "131");
