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
    and_false_left_no_eval,
    "int x = 0; if (false && (x = 1) == 1) {} System.out.println(x);",
    "0"
);
jt!(
    and_true_left_evaluates_right,
    "int x = 0; if (true && (x = 2) == 2) {} System.out.println(x);",
    "2"
);
jt!(
    or_true_left_no_eval,
    "int x = 0; if (true || (x = 3) == 3) {} System.out.println(x);",
    "0"
);
jt!(
    or_false_left_evaluates_right,
    "int x = 0; if (false || (x = 4) == 4) {} System.out.println(x);",
    "4"
);
jt!(
    and_chain_no_eval_second,
    "int x = 0; int y = 0; if ((x == 1) && ((x = 2) == 2) && ((y = 3) == 3)) {} System.out.println(x + y);",
    "0"
);
jt!(
    and_chain_evaluates_second,
    "int x = 1; int y = 0; if ((x == 1) && ((x = 2) == 2) && ((y = 3) == 3)) {} System.out.println(x + y);",
    "5"
);
jt!(
    or_chain_no_eval_second,
    "int x = 0; int y = 0; if ((x == 0) || ((x = 2) == 2) || ((y = 1) == 1)) {} System.out.println(x + y);",
    "0"
);
jt!(
    or_chain_evaluates_second,
    "int x = 1; int y = 0; if ((x == 0) || ((x = 2) == 2) || ((y = 1) == 1)) {} System.out.println(x + y);",
    "2"
);
jt!(not_true, "System.out.println(!false);", "true");
jt!(not_false, "System.out.println(!true);", "false");
jt!(
    not_preserves_and,
    "int x = 0; System.out.println(!(false && (x = 1) == 1) ? x : 1);",
    "0"
);
jt!(
    logical_precedence_1,
    "int x = 0; System.out.println(true || false && false);",
    "true"
);
jt!(
    logical_precedence_2,
    "int x = 0; System.out.println((true || false) && false);",
    "false"
);
jt!(
    while_with_and,
    "int x = 0; int i = 0; while (i < 4 && i < 3) { x += i; i++; } System.out.println(x);",
    "3"
);
jt!(
    while_with_or,
    "int x = 0; int i = 0; while (false || i < 3) { x += i; i++; if (i == 3) break; } System.out.println(x);",
    "3"
);
jt!(
    do_while_with_and,
    "int x = 0; int i = 0; do { if ((i > 1) && (x = 5) == 5) { break; } i++; } while (i < 3); System.out.println(x);",
    "0"
);
jt!(
    and_skip_with_assignment,
    "int x = 0; if ((x > 0) && (x = 1) == 1) { x = 2; } System.out.println(x);",
    "0"
);
jt!(
    or_skip_with_assignment,
    "int x = 0; if ((x == 0) || (x = 1) == 1) { x = 3; } System.out.println(x);",
    "3"
);
jt!(
    short_circuit_in_while,
    "int x = 0; int i = 0; while (false && (i++ > 0)) { x++; } System.out.println(i + x);",
    "0"
);
jt!(
    no_short_circuit_and,
    "int x = 0; if (false & (x = 1) == 1) {} System.out.println(x);",
    "1"
);
jt!(
    no_short_circuit_or,
    "int x = 0; if (true | (x = 2) == 2) {} System.out.println(x);",
    "2"
);
jt!(
    bitwise_vs_logical_and,
    "int x = 0; int y = (0 & 1) | 2; System.out.println(y);",
    "2"
);
jt!(
    bitwise_vs_logical_or,
    "int x = 0; int y = (1 | 2); System.out.println(y);",
    "3"
);
jt!(xor_logical, "System.out.println(true ^ false);", "true");
jt!(xor_false, "System.out.println(false ^ false);", "false");
jt!(
    complex_short_circuit_1,
    "int x = 0; int y = (false && (x = 1) == 1) || (true && (x = 2) == 2) ? x : x; System.out.println(y);",
    "2"
);
jt!(
    complex_short_circuit_2,
    "int x = 0; int y = (x == 1 && (x = 2) == 2) || (x == 0 && (x = 3) == 3) ? x : x; System.out.println(y);",
    "3"
);
