use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(unary_plus_identity, "System.out.println(+5);", "5");
jt!(unary_minus, "System.out.println(-5);", "-5");
jt!(unary_plus_zero, "System.out.println(+0);", "0");
jt!(double_negation, "System.out.println(-(-7));", "7");
jt!(prefix_increment, "int x = 3; System.out.println(++x);", "4");
jt!(
    prefix_increment_then_print,
    "int x = 3; ++x; System.out.println(x);",
    "4"
);
jt!(
    postfix_increment,
    "int x = 3; System.out.println(x++);",
    "3"
);
jt!(
    postfix_increment_followed,
    "int x = 3; x++; System.out.println(x);",
    "4"
);
jt!(prefix_decrement, "int x = 3; System.out.println(--x);", "2");
jt!(
    prefix_decrement_then_print,
    "int x = 3; --x; System.out.println(x);",
    "2"
);
jt!(
    postfix_decrement,
    "int x = 3; System.out.println(x--);",
    "3"
);
jt!(
    postfix_decrement_followed,
    "int x = 3; x--; System.out.println(x);",
    "2"
);
jt!(bitwise_not, "System.out.println(~0);", "-1");
jt!(bitwise_not_nonzero, "System.out.println(~1);", "-2");
jt!(bitwise_not_negative, "System.out.println(~(-2));", "1");
jt!(
    pre_inc_nested,
    "int x = 1; System.out.println(+(++x));",
    "2"
);
jt!(
    post_inc_in_expr,
    "int x = 1; System.out.println(x++ + 2);",
    "3"
);
jt!(
    post_inc_after_expr,
    "int x = 1; System.out.println(2 + x++);",
    "3"
);
jt!(
    post_inc_then_post_inc,
    "int x = 1; int y = x++ + x++; System.out.println(y);",
    "3"
);
jt!(
    inc_preserves_reference,
    "int[] box = {1}; box[0]++; System.out.println(box[0]);",
    "2"
);
jt!(
    dec_preserves_reference,
    "int[] box = {1}; box[0]--; System.out.println(box[0]);",
    "0"
);
jt!(
    unary_on_reference,
    "int[] a = {1}; System.out.println(a[0]);",
    "1"
);
jt!(
    multiple_unary_mixed,
    "int[] arr = {5}; System.out.println(-arr[0]);",
    "-5"
);
jt!(
    binary_plus_then_unary,
    "System.out.println(-(2 + 3));",
    "-5"
);
jt!(
    unary_then_multiplicative,
    "System.out.println(-2 * 3);",
    "-6"
);
jt!(
    plus_with_boolean_not,
    "System.out.println(!(!false));",
    "true"
);
