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
    nested_for_depth_two,
    "int sum = 0; for(int i = 0; i < 3; i++) { for(int j = 0; j < 2; j++) sum++; } System.out.println(sum);",
    "6"
);
jt!(
    nested_while_with_for,
    "int sum = 0; int i = 0; while(i < 2) { for(int j = 0; j < 2; j++) sum++; i++; } System.out.println(sum);",
    "4"
);
jt!(
    three_level_triple_nested,
    "int c = 0; for(int i = 0; i < 2; i++) { for(int j = 0; j < 2; j++) { for(int k = 0; k < 2; k++) c++; } } System.out.println(c);",
    "8"
);
jt!(
    nested_break_inner_only,
    "int i = 0; int j = 0; int c = 0; for(; i < 3; i++) { for(; j < 3; j++) { if(j == 1) break; c++; } j = 0; } System.out.println(c);",
    "3"
);
jt!(
    nested_continue_inner,
    "int i = 0; int c = 0; for(; i < 3; i++) { int j = 0; while(j < 3) { j++; if(j == 2) continue; c++; } } System.out.println(c);",
    "6"
);
jt!(
    nested_break_outer_by_condition,
    "int i = 0; int c = 0; outer: for(; i < 5; i++) { for(int j = 0; j < 5; j++) { if(i == 2) break outer; c++; } } System.out.println(c);",
    "10"
);
jt!(
    nested_label_continue,
    "int c = 0; outer: for(int i = 0; i < 3; i++) { for(int j = 0; j < 3; j++) { if(j == 1) continue outer; c++; } } System.out.println(c);",
    "3"
);
jt!(
    nested_boolean_guard,
    "int c = 0; for(int i = 0; i < 3; i++) { if(i == 1) { for(int j = 0; j < 3; j++) c++; } else { c += 2; } } System.out.println(c);",
    "7"
);
jt!(
    nested_sum_matrix_like,
    "int total = 0; for(int i = 0; i < 3; i++) { int row = i * 10; for(int j = 0; j < 2; j++) total += row + j; } System.out.println(total);",
    "63"
);
jt!(
    nested_with_mutating_outer_counter,
    "int total = 0; int i = 0; for(; i < 3; i++) { int j = 0; while(j < 2) { if(j == 0) i++; total++; j++; } } System.out.println(total + i);",
    "8"
);
jt!(
    nested_empty_body,
    "int c = 0; for(int i = 0; i < 2; i++) { for(int j = 0; j < 2; j++) { } c++; } System.out.println(c);",
    "2"
);
jt!(
    nested_with_array_rows,
    "int[][] m = {{1,2},{3,4},{5,6}}; int s = 0; for(int i = 0; i < m.length; i++) { for(int j = 0; j < m[i].length; j++) s += m[i][j]; } System.out.println(s);",
    "21"
);
jt!(
    nested_string_concat,
    "int c = 0; for(int i = 0; i < 2; i++) { for(int j = 0; j < 2; j++) c++; } System.out.println(c == 4);",
    "true"
);
jt!(
    while_inside_for,
    "int c = 0; for(int i = 0; i < 3; i++) { int j = 0; while(j < 2) { c += j; j++; } } System.out.println(c);",
    "3"
);
jt!(
    do_while_inside_for,
    "int c = 0; for(int i = 0; i < 2; i++) { int j = 0; do { c++; j++; } while(j < 2); } System.out.println(c);",
    "4"
);
jt!(
    nested_break_two_levels,
    "int c = 0; outer: for(int i = 0; i < 5; i++) { for(int j = 0; j < 5; j++) { if(i + j == 3) break outer; c++; } } System.out.println(c);",
    "3"
);
jt!(
    nested_break_inner_level_only,
    "int c = 0; for(int i = 0; i < 5; i++) { for(int j = 0; j < 5; j++) { if(i == 3) break; c++; } } System.out.println(c);",
    "20"
);
jt!(
    nested_continue_reaches_outer,
    "int c = 0; for(int i = 0; i < 3; i++) { for(int j = 0; j < 3; j++) { if(j == 1) continue; c++; } } System.out.println(c);",
    "6"
);
jt!(
    nested_continue_outer_label,
    "int c = 0; outer: for(int i = 0; i < 3; i++) { for(int j = 0; j < 3; j++) { if(j == 1) continue outer; c++; } } System.out.println(c);",
    "3"
);
jt!(
    nested_if_else,
    "int c = 0; for(int i = 0; i < 4; i++) { for(int j = 0; j < 2; j++) { if((i + j) % 2 == 0) c++; else c += 2; } } System.out.println(c);",
    "12"
);
jt!(
    nested_break_condition,
    "int c = 0; for(int i = 0; i < 4; i++) { for(int j = 0; j < 4; j++) { if(i == 2 && j == 2) break; c++; } } System.out.println(c);",
    "14"
);
jt!(
    nested_while_break,
    "int c = 0; for(int i = 0; i < 2; i++) { int j = 0; while(j < 5) { if(j == 3) break; c++; j++; } } System.out.println(c);",
    "6"
);
jt!(
    nested_while_continue,
    "int c = 0; for(int i = 0; i < 2; i++) { int j = 0; while(j < 5) { j++; if(j % 2 == 0) continue; c++; } } System.out.println(c);",
    "6"
);
jt!(
    nested_do_while_break,
    "int c = 0; for(int i = 0; i < 2; i++) { int j = 0; do { if(j == 2) break; c++; j++; } while(true); } System.out.println(c);",
    "4"
);
jt!(
    mixed_nested_constructs,
    "int total = 0; for(int i = 0; i < 2; i++) { int j = 0; while(j < 3) { for(int k = 0; k < 2; k++) total += i + j + k; j++; } } System.out.println(total);",
    "24"
);
jt!(
    nested_boolean_guarding,
    "int total = 0; for(int i = 0; i < 3; i++) { boolean ready = i % 2 == 0; for(int j = 0; j < 2; j++) { if(!ready) continue; total += j; } } System.out.println(total);",
    "2"
);
jt!(
    nested_with_variable_reshadow,
    "int total = 0; for(int i = 0; i < 2; i++) { for(int j = 0; j < 2; j++) { int totalInner = i + j; total += totalInner; } } System.out.println(total);",
    "4"
);
