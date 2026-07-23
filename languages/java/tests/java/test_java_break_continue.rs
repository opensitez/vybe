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
    break_exits_while,
    "int i = 0; while(true) { if(i == 3) break; i++; } System.out.println(i);",
    "3"
);
jt!(
    break_while_after_increment,
    "int i = 0; while(i < 5) { i++; if(i == 2) break; } System.out.println(i);",
    "2"
);
jt!(
    continue_skips_incremented_body,
    "int i = 0; int s = 0; for(i = 0; i < 4; i++) { if(i == 1) continue; s += i; } System.out.println(s);",
    "6"
);
jt!(
    continue_in_nested,
    "int c = 0; for(int i = 0; i < 3; i++) { for(int j = 0; j < 3; j++) { if(j == 1) continue; c++; } } System.out.println(c);",
    "6"
);
jt!(
    break_nested_inner,
    "int c = 0; for(int i = 0; i < 3; i++) { for(int j = 0; j < 3; j++) { if(j == 2) break; c++; } } System.out.println(c);",
    "6"
);
jt!(
    break_nested_outer_label,
    "int c = 0; outer: for(int i = 0; i < 3; i++) { for(int j = 0; j < 3; j++) { if(i == 1) break outer; c++; } } System.out.println(c);",
    "3"
);
jt!(
    continue_nested_outer_label,
    "int c = 0; outer: for(int i = 0; i < 3; i++) { for(int j = 0; j < 3; j++) { if(j == 1) continue outer; c++; } } System.out.println(c);",
    "0"
);
jt!(
    break_after_do,
    "int i = 0; do { if(i == 2) break; i++; } while(true); System.out.println(i);",
    "2"
);
jt!(
    continue_after_do,
    "int i = 0; int s = 0; do { i++; if(i % 2 == 0) continue; s++; } while(i < 5); System.out.println(s);",
    "3"
);
jt!(
    break_before_continue,
    "int i = 0; int s = 0; while(i < 5) { i++; if(i == 2) { break; } s++; } System.out.println(s);",
    "1"
);
jt!(
    labelled_continue_while,
    "int total = 0; int i = 0; outer: while(i < 3) { i++; int j = 0; while(j < 3) { j++; if(j == 2) continue outer; total++; } } System.out.println(total);",
    "0"
);
jt!(
    labelled_break_from_switch_like,
    "int total = 0; outer: for(int i = 0; i < 2; i++) { for(int j = 0; j < 2; j++) { if(j == 1) break outer; total++; } } System.out.println(total);",
    "0"
);
jt!(
    break_with_if_guard,
    "int i = 0; int s = 0; while(i < 10) { if(i == 4) break; if(i % 2 == 0) s += i; i++; } System.out.println(s);",
    "6"
);
jt!(
    continue_with_assignment,
    "int i = 0; int s = 0; for(i = 0; i < 5; i++) { if(i == 3) continue; s += 1; } System.out.println(s);",
    "4"
);
jt!(
    break_then_else_path,
    "int i = 0; int s = 0; while(true) { if(i == 2) break; if(i == 1) { i++; continue; } s += i; i++; } System.out.println(s);",
    "1"
);
jt!(
    continue_no_ops,
    "int i = 0; int c = 0; for(; i < 4; i++) { if(i == 2) continue; c++; } System.out.println(c);",
    "3"
);
jt!(
    break_two_level_condition,
    "int c = 0; for(int i = 0; i < 4; i++) { for(int j = 0; j < 4; j++) { if(i + j > 3) break; c++; } } System.out.println(c);",
    "10"
);
jt!(
    continue_two_level_condition,
    "int c = 0; for(int i = 0; i < 3; i++) { for(int j = 0; j < 3; j++) { if((i + j) % 2 == 0) continue; c++; } } System.out.println(c);",
    "3"
);
jt!(
    break_from_if_inside_for,
    "int c = 0; for(int i = 0; i < 5; i++) { if(i == 3) break; c++; } System.out.println(c);",
    "3"
);
jt!(
    continue_outside_if,
    "int c = 0; for(int i = 0; i < 5; i++) { if(i < 3) continue; c++; } System.out.println(c);",
    "2"
);
jt!(
    break_zero_iterations,
    "int c = 0; for(int i = 0; i < 3; i++) { if(i > 10) break; c++; } System.out.println(c);",
    "3"
);
jt!(
    continue_zero,
    "int c = 0; for(int i = 0; i < 3; i++) { if(i < 0) continue; c++; } System.out.println(c);",
    "3"
);
jt!(
    label_without_jump_no_effect,
    "int c = 0; outer: for(int i = 0; i < 3; i++) { c++; } System.out.println(c);",
    "3"
);
jt!(
    nested_labels_untaken,
    "int c = 0; outer: for(int i = 0; i < 2; i++) { inner: for(int j = 0; j < 2; j++) { c++; } } System.out.println(c);",
    "4"
);
jt!(
    break_with_return_pattern,
    "int c = 0; for(int i = 0; i < 5; i++) { if(i == 3) { c = i; break; } } System.out.println(c);",
    "3"
);
jt!(
    continue_with_expression,
    "int c = 0; for(int i = 0; i < 5; i++) { if(i == 4) continue; c += i + 1; } System.out.println(c);",
    "10"
);
