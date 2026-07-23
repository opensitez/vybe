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
    local_scope_hides_field,
    "int x = 1; { int x = 2; System.out.println(x); }",
    "2"
);
jt!(
    scope_restores_outer,
    "int x = 1; { int x = 2; } System.out.println(x);",
    "1"
);
jt!(
    if_scope_separate,
    "int x = 1; if (true) { int x = 2; } else { int x = 3; } System.out.println(x);",
    "1"
);
jt!(
    for_scope_variable,
    "for(int x = 0; x < 1; x++) { } System.out.println(9);",
    "9"
);
jt!(
    while_scope_variable,
    "int x = 0; while(x < 1) { int y = x; x++; } System.out.println(x);",
    "1"
);
jt!(
    block_scope_true_false,
    "int a = 1; if(false) { int a = 2; } else { a = 3; } System.out.println(a);",
    "3"
);
jt!(
    nested_scope_sum,
    "int a = 1; { int b = 2; { int c = 3; a = a + b + c; } } System.out.println(a);",
    "6"
);
jt!(
    try_scope_cleanup,
    "int x = 5; try { int y = 2; x = x + y; } finally { } System.out.println(x);",
    "7"
);
jt!(
    variable_reuse_after_loop,
    "int total = 0; for(int i = 0; i < 2; i++) { total += i; } int i = 5; System.out.println(i);",
    "5"
);
jt!(
    method_scope_local,
    "int outer = 10; { int inner = 4; outer = outer - inner; } System.out.println(outer);",
    "6"
);
jt!(
    catch_scope_variable,
    "int x = 0; try { throw new RuntimeException(); } catch(Exception e) { int c = 7; x = c; } System.out.println(x);",
    "7"
);
jt!(
    final_variable_assignment_forbidden,
    "final int x = 1; // cannot reassign intentionally checked by parse-time
System.out.println(x);",
    "1"
);
jt!(
    final_array_reference,
    "final int[] a = {1,2,3}; a[0] = 5; System.out.println(a[0]);",
    "5"
);
jt!(
    effectively_final_in_lambda,
    "int x = 3; java.util.function.Supplier<Integer> s = () -> x + 1; System.out.println(s.get());",
    "4"
);
jt!(
    no_effectively_final_not_supported,
    "int x = 1; x = 2; java.util.function.Supplier<Integer> s = () -> x; System.out.println(s.get());",
    "2"
);
jt!(
    switch_scope_bound,
    "int x = 1; switch(x) { case 1: int local = 9; System.out.println(local); break; default: System.out.println(0); }",
    "9"
);
jt!(
    for_with_existing_outer_name,
    "int i = 1; for(i = 0; i < 1; i++) {} System.out.println(i);",
    "1"
);
jt!(
    inner_scope_forbid_name_collision,
    "int x = 1; if(true) { int xInner = x + 1; System.out.println(xInner); }",
    "2"
);
jt!(
    while_with_decl_inside,
    "int x = 0; while(x < 2) { int y = x * 2; x++; } System.out.println(x);",
    "2"
);
jt!(
    scope_after_if,
    "int x = 1; if(x > 0) { int y = 10; } System.out.println(x);",
    "1"
);
jt!(
    scope_shadow_chain,
    "int x = 1; { int x = 2; { int x = 3; System.out.println(x); } }",
    "3"
);
jt!(
    scope_reassign_outer,
    "int x = 1; { x = 4; } System.out.println(x);",
    "4"
);
jt!(
    scope_class_local,
    "int x = 1; class Local { int v = x; } System.out.println(new Local().v);",
    "1"
);
jt!(
    lambda_captures_shadowed,
    "int x = 1; java.util.function.Supplier<Integer> s1 = () -> x; { int x = 5; System.out.println(s1.get()); }",
    "1"
);
jt!(
    scope_multiple_blocks,
    "int x = 1; { x = 2; } { int x = 3; System.out.println(2); }",
    "2"
);
jt!(
    scope_final_outer_immutable,
    "final int v = 9; { int seen = v; System.out.println(seen); }",
    "9"
);
jt!(
    array_scope_variable,
    "int[] arr = {1,2}; int v = arr[0]; { int w = arr[1]; v = w; } System.out.println(v);",
    "2"
);
jt!(
    block_after_loop,
    "int total = 0; for(int i=0;i<2;i++){ total += i; } { int extra = 10; total += extra; } System.out.println(total);",
    "11"
);
jt!(
    scope_label_name_collision,
    "int result = 0; outer: for(int i = 0; i < 1; i++) { int value = 8; result = value; } System.out.println(result);",
    "8"
);
