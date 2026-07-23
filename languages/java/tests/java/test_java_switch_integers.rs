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
    single_case,
    "int n = 2; int v = 0; switch (n) { case 1: v = 10; break; case 2: v = 20; break; default: v = 30; } System.out.println(v);",
    "20"
);
jt!(
    default_case,
    "int n = 5; int v = 0; switch (n) { case 1: v = 10; break; case 2: v = 20; break; default: v = 30; } System.out.println(v);",
    "30"
);
jt!(
    multiple_cases,
    "int n = 3; int v = 0; switch (n) { case 1: v = 1; break; case 2: v = 2; break; case 3: v = 3; break; default: v = 0; } System.out.println(v);",
    "3"
);
jt!(
    switch_on_expression,
    "int n = 4; int v = 0; switch (n / 2) { case 1: v = 5; break; case 2: v = 6; break; default: v = 0; } System.out.println(v);",
    "6"
);
jt!(
    grouped_cases,
    "int n = 3; int v = 0; switch (n) { case 1: case 2: v = 1; break; case 3: case 4: v = 2; break; default: v = 0; } System.out.println(v);",
    "2"
);
jt!(
    string_from_switch,
    "int n = 4; String s = \"\"; switch (n) { case 4: s = \"four\"; break; case 5: s = \"five\"; break; default: s = \"other\"; } System.out.println(s);",
    "four"
);
jt!(
    switch_with_fall_to_default,
    "int n = 4; int v = 0; switch (n) { case 1: v = 10; break; case 2: case 3: v = 20; break; case 4: v = 30; default: v = 40; } System.out.println(v);",
    "40"
);
jt!(
    switch_updates_array,
    "int n = 1; int[] a = {0,1,2,3}; switch(n){ case 0: a[0]=9; break; case 1: a[1]=9; break; default: a[2]=9; } System.out.println(a[1]);",
    "9"
);
jt!(
    switch_return_emulation,
    "int n = 2; int v = 0; switch (n) { case 1: v = 5; break; case 2: v = 10; break; case 3: v = 15; break; } System.out.println(v);",
    "10"
);
jt!(
    switch_composed_condition,
    "int n = 7; int v = 0; switch (n % 3) { case 0: v = 3; break; case 1: v = 1; break; case 2: v = 2; break; } System.out.println(v);",
    "1"
);
jt!(
    switch_in_loop,
    "int sum = 0; for (int i = 0; i < 4; i++) { int v = 0; switch (i) { case 0: v = 1; break; case 1: v = 2; break; case 2: v = 3; break; default: v = 4; } sum += v; } System.out.println(sum);",
    "10"
);
jt!(
    nested_switch,
    "int a = 1; int b = 2; int v = 0; switch (a) { case 1: switch (b) { case 2: v = 12; break; default: v = 10; } break; default: v = 0; } System.out.println(v);",
    "12"
);
jt!(
    switch_on_byte_like,
    "byte n = 1; int v = 0; switch (n) { case 1: v = 11; break; case 2: v = 22; break; default: v = 33; } System.out.println(v);",
    "11"
);
jt!(
    switch_with_math,
    "int x = 8; int y = 0; switch (x) { case 2: y = 2; break; case 4: y = 4; break; case 8: y = 8; break; default: y = 0; } System.out.println(y + x / 2);",
    "12"
);
jt!(
    switch_zero,
    "int n = 0; int v = 0; switch (n) { case 0: v = 0; break; default: v = 1; } System.out.println(v);",
    "0"
);
jt!(
    switch_negative,
    "int n = -1; int v = 0; switch (n) { case -1: v = 9; break; case 1: v = 1; break; default: v = 0; } System.out.println(v);",
    "9"
);
jt!(
    switch_with_chars,
    "int c = 'b'; int v = 0; switch (c) { case 'a': v = 1; break; case 'b': v = 2; break; default: v = 0; } System.out.println(v);",
    "2"
);
jt!(
    switch_with_string_builder,
    "int n = 6; String s = \"\"; switch (n) { case 5: s += \"a\"; break; case 6: s += \"b\"; break; default: s += \"x\"; } System.out.println(s);",
    "b"
);
jt!(
    switch_many_cases,
    "int n = 5; int v = 0; switch (n) { case 1: v=1; break; case 2: v=2; break; case 3: v=3; break; case 4: v=4; break; case 5: v=5; break; default: v=0; } System.out.println(v);",
    "5"
);
jt!(
    switch_boolean_result,
    "int n = 2; boolean b = false; switch (n) { case 1: b = false; break; case 2: b = true; break; default: b = false; } System.out.println(b);",
    "true"
);
jt!(
    switch_in_expression_order,
    "int n = 2; int v = 0; switch (n - 1) { case 1: v = 3; break; case 0: v = 2; break; default: v = 1; } System.out.println(v);",
    "3"
);
jt!(
    switch_with_complex_default,
    "int n = 42; int v = 0; switch (n) { case 1: v = 1; break; case 2: v = 2; break; default: v = 40; } System.out.println(v - 2);",
    "38"
);
jt!(
    switch_constant,
    "int n = 7; final int A = 7; int v = 0; switch (n) { case A: v = 70; break; default: v = 0; } System.out.println(v);",
    "70"
);
jt!(
    switch_no_break_chain,
    "int n = 1; int v = 0; switch (n) { case 1: v += 1; case 2: v += 2; case 3: v += 3; default: v += 4; } System.out.println(v);",
    "10"
);
jt!(
    switch_with_block,
    "int n = 2; int v = 0; switch (n) { case 1: { v = 1; } break; case 2: { v = 2; } break; default: { v = 0; } } System.out.println(v);",
    "2"
);
jt!(
    switch_after_loop,
    "int n = 3; int v = 0; for (int i = 0; i < 1; i++) { switch (n) { case 1: v = 1; break; case 2: v = 2; break; case 3: v = 3; break; } } System.out.println(v);",
    "3"
);
