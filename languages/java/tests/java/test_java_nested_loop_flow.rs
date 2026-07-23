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
    nested_addition,
    "int s = 0; for (int i = 0; i < 3; i++) { for (int j = 0; j < 2; j++) { s += i + j; } } System.out.println(s);",
    "9"
);
jt!(
    nested_multiplication,
    "int s = 1; for (int i = 1; i <= 3; i++) { for (int j = 1; j <= 2; j++) { s *= j; } } System.out.println(s);",
    "8"
);
jt!(
    nested_with_continue,
    "int s = 0; for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { if (j == 1) continue; s += i + j; } } System.out.println(s);",
    "12"
);
jt!(
    nested_with_break,
    "int s = 0; for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { if (i + j == 3) break; s += 1; } } System.out.println(s);",
    "6"
);
jt!(
    nested_while_for,
    "int s = 0; int i = 0; while (i < 2) { for (int j = 0; j < 3; j++) { s += i + j; } i++; } System.out.println(s);",
    "9"
);
jt!(
    nested_for_while,
    "int s = 0; for (int i = 0; i < 2; i++) { int j = 0; while (j < 2) { s += i * j; j++; } } System.out.println(s);",
    "1"
);
jt!(
    inner_increments_outer,
    "int s = 0; for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { s += j; i += (j == 2 ? 0 : 0); } } System.out.println(s);",
    "9"
);
jt!(
    matrix_row_sum,
    "int[][] m = {{1,2},{3,4}}; int s = 0; for (int i = 0; i < m.length; i++) { for (int j = 0; j < m[i].length; j++) { s += m[i][j]; } } System.out.println(s);",
    "10"
);
jt!(
    matrix_find,
    "int[][] m = {{1,2},{3,4}}; int hit = 0; for (int i = 0; i < m.length; i++) { for (int j = 0; j < m[i].length; j++) { if (m[i][j] == 3) hit = i * 10 + j; } } System.out.println(hit);",
    "10"
);
jt!(
    nested_if_else,
    "int s = 0; for (int i = 0; i < 4; i++) { for (int j = 0; j < 4; j++) { if ((i + j) % 2 == 0) s += 1; else s += 2; } } System.out.println(s);",
    "24"
);
jt!(
    nested_boolean_guard,
    "int s = 0; for (int i = 0; i < 4; i++) { if (i > 1) for (int j = 0; j < 2; j++) { s += j; } else s += 5; } System.out.println(s);",
    "12"
);
jt!(
    nested_local_reset,
    "int total = 0; for (int i = 0; i < 3; i++) { int row = 0; for (int j = 0; j < 2; j++) { row += j; } total += row; } System.out.println(total);",
    "3"
);
jt!(
    nested_depth_two,
    "int total = 0; for (int i = 0; i < 2; i++) { for (int j = 0; j < 2; j++) { for (int k = 0; k < 2; k++) { total += 1; } } } System.out.println(total);",
    "8"
);
jt!(
    nested_depth_three,
    "int total = 0; for (int i = 0; i < 2; i++) { for (int j = 0; j < 2; j++) { for (int k = 0; k < 2; k++) { total += i + j + k; } } } System.out.println(total);",
    "12"
);
jt!(
    nested_with_labeled_like,
    "int total = 0; for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { if (i == 2 && j == 2) continue; total += i * j; } } System.out.println(total);",
    "5"
);
jt!(
    nested_array_lookup,
    "int[][] m = {{0,1},{2,3},{4,5}}; int n = 0; for (int i = 0; i < m.length; i++) { for (int j = 0; j < m[i].length; j++) { if (m[i][j] > n) n = m[i][j]; } } System.out.println(n);",
    "5"
);
jt!(
    nested_composition,
    "int total = 0; for (int i = 1; i <= 3; i++) { for (int j = 1; j <= 2; j++) { total += i * j; } } System.out.println(total);",
    "18"
);
jt!(
    nested_even_outer,
    "int total = 0; for (int i = 0; i < 5; i++) { if ((i & 1) == 0) { for (int j = 0; j < 2; j++) { total += i; } } } System.out.println(total);",
    "12"
);
jt!(
    nested_odd_outer,
    "int total = 0; for (int i = 0; i < 5; i++) { for (int j = 0; j < 2; j++) { total += (i & 1); } } System.out.println(total);",
    "4"
);
jt!(
    nested_char_matrix,
    "char[][] m = {{'a','b'},{'c'}}; int total = 0; for (int i = 0; i < m.length; i++) { for (int j = 0; j < m[i].length; j++) { total += m[i][j] - 96; } } System.out.println(total);",
    "6"
);
jt!(
    nested_string_matrix,
    "String[][] s = {{\"x\",\"y\"},{\"z\"}}; int total = 0; for (int i = 0; i < s.length; i++) { for (int j = 0; j < s[i].length; j++) { total += s[i][j].length(); } } System.out.println(total);",
    "3"
);
jt!(
    nested_sum_break,
    "int total = 0; for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { total += 1; if (total == 5) break; } } System.out.println(total);",
    "8"
);
jt!(
    nested_sum_conditional_continue,
    "int total = 0; for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { if ((i + j) == 2) continue; total += 1; } } System.out.println(total);",
    "6"
);
jt!(
    nested_with_flags,
    "int ones = 0; boolean active = true; for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { if (active && (i + j) == 3) active = false; if (active) ones++; } } System.out.println(ones);",
    "5"
);
jt!(
    nested_after_break_outer,
    "int total = 0; for (int i = 0; i < 3; i++) { for (int j = 0; j < 3; j++) { total++; if (j == 1) break; } if (i == 1) break; } System.out.println(total);",
    "4"
);
jt!(
    nested_while_in_for,
    "int total = 0; for (int i = 0; i < 3; i++) { int j = 0; while (j <= i) { total += j; j++; } } System.out.println(total);",
    "4"
);
