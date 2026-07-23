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
    for_loop_sum_to_three,
    "int s = 0; for(int i = 1; i <= 3; i++) { s += i; } System.out.println(s);",
    "6"
);
jt!(
    for_loop_init_expression,
    "int n = 0; for(int i = 0; i < 4; i = i + 1) n += i; System.out.println(n);",
    "6"
);
jt!(
    for_loop_nested_increment,
    "int c = 0; for(int i = 0; i < 2; i++) { for(int j = 0; j < 2; j++) { c++; } } System.out.println(c);",
    "4"
);
jt!(
    for_loop_no_init,
    "int i = 0; int c = 0; for(; i < 3; i++) c++; System.out.println(c);",
    "3"
);
jt!(
    for_loop_no_update,
    "int c = 0; for(int i = 0; i < 1; ) { c++; i++; } System.out.println(c);",
    "1"
);
jt!(
    while_loop_count,
    "int c = 0; int i = 1; while(i <= 3) { c += i; i++; } System.out.println(c);",
    "6"
);
jt!(
    while_loop_no_exec,
    "int c = 0; int i = 5; while(i < 1) c++; System.out.println(c);",
    "0"
);
jt!(
    do_while_execs_once,
    "int c = 0; int i = 0; do { c++; i++; } while(i > 10); System.out.println(c);",
    "1"
);
jt!(
    do_while_conditional_loop,
    "int c = 0; int i = 0; do { c++; i++; } while(i < 3); System.out.println(c);",
    "3"
);
jt!(
    while_with_break,
    "int i = 0; while(i < 10) { if (i == 4) break; i++; } System.out.println(i);",
    "4"
);
jt!(
    while_with_continue,
    "int i = 0; int s = 0; while(i < 5) { i++; if (i % 2 == 0) continue; s++; } System.out.println(s);",
    "3"
);
jt!(
    for_break_first,
    "int c = 0; for(int i = 0; i < 10; i++) { if(i == 2) break; c++; } System.out.println(c);",
    "2"
);
jt!(
    for_continue_first,
    "int c = 0; for(int i = 0; i < 5; i++) { if(i == 3) continue; c++; } System.out.println(c);",
    "4"
);
jt!(
    for_else_like_flow,
    "int c = 0; for(int i = 0; i < 4; i++) { if(i == 1) { continue; } c++; } System.out.println(c);",
    "3"
);
jt!(
    while_multiple_updates,
    "int i = 0; int a = 0; while(i < 4) { a += i; i++; } System.out.println(a);",
    "6"
);
jt!(
    while_with_assignment,
    "int i = 1; while(i <= 4) { i = i + 1; } System.out.println(i);",
    "5"
);
jt!(
    do_while_with_mod,
    "int i = 0; int s = 0; do { s += i; i++; } while(i < 4); System.out.println(s);",
    "6"
);
jt!(
    loop_variable_scope,
    "int total = 0; for(int i = 0; i < 3; i++) { int local = 1; total += local; } System.out.println(total);",
    "3"
);
jt!(
    empty_loop_body,
    "int i = 0; for(; i < 3; i++) { } System.out.println(i);",
    "3"
);
jt!(
    while_empty,
    "int i = 0; while(i < 1) { i++; } System.out.println(i);",
    "1"
);
jt!(
    for_update_more_complex,
    "int i = 1; int c = 0; for(; i <= 4; i = i + 2) c++; System.out.println(c);",
    "2"
);
jt!(
    for_with_decl_and_postfix,
    "int c = 0; for(int i = 2; i > 0; i--) c += i; System.out.println(c);",
    "3"
);
jt!(
    nested_while,
    "int i = 0; int c = 0; while(i < 2) { int j = 0; while(j < 2){ c++; j++; } i++; } System.out.println(c);",
    "4"
);
jt!(
    while_if_inside,
    "int i = 0; int s = 0; while(i < 5) { if(i % 2 == 0) s += i; i++; } System.out.println(s);",
    "6"
);
jt!(
    for_condition_with_method_call,
    "int[] arr = {1,2,3}; int s = 0; for(int i = 0; i < arr.length; i++) s += arr[i]; System.out.println(s);",
    "6"
);
jt!(
    loop_reassigned_counter,
    "int i = 0; while(true) { if(i == 2) break; i = i + 1; } System.out.println(i);",
    "2"
);
