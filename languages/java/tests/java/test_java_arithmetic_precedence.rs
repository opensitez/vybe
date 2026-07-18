use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(add_before_multiply, "System.out.println(2 + 3 * 4);", "14");
jt!(multiply_before_subtract, "System.out.println(20 - 3 * 2);", "14");
jt!(divide_before_add, "System.out.println(20 / 2 + 3);", "13");
jt!(add_before_divide, "System.out.println(20 + 7 / 2);", "23");
jt!(parenthesized_add_first, "System.out.println((2 + 3) * 4);", "20");
jt!(parenthesized_divide_first, "System.out.println(20 / (2 + 3));", "4");
jt!(modulo_vs_add, "System.out.println(13 % 3 + 1);", "2");
jt!(add_vs_bitwise_or, "System.out.println(1 + 2 | 4);", "7");
jt!(bitwise_first, "System.out.println((1 + 2) | 4);", "7");
jt!(comparison_precedes_bitwise, "System.out.println(1 + 2 > 2 ? 10 : 1);", "10");
jt!(and_vs_or_precedence, "System.out.println(1 + 2 | 4 & 3);", "7");
jt!(bitwise_and_precedes_or, "System.out.println(1 + (2 | 3));", "6");
jt!(shift_precedence, "System.out.println(1 + 1 << 2);", "4");
jt!(shift_with_parentheses, "System.out.println((1 + 1) << 2);", "8");
jt!(nested_arithmetic, "System.out.println(10 - (3 * (2 + 1)) + 5);", "12");
jt!(multiple_levels, "System.out.println((2 + 3) * (4 - 1) / 5);", "3");
jt!(unary_precedence, "System.out.println(-3 + 10);", "7");
jt!(unary_and_precedence, "System.out.println(-(3 + 2));", "-5");
jt!(nested_conditional, "System.out.println((2 > 1 ? 10 : 20) + 1);", "11");
jt!(equality_chain_left_assoc, "System.out.println(1 + 2 == 3);", "true");
jt!(equality_and_relational, "System.out.println(2 * 3 == 6);", "true");
jt!(arithmetic_and_equality, "System.out.println((2 + 2) * 2 == 8);", "false");
jt!(complex_expression, "System.out.println(1 + 2 + 3 * 4 / 2 - 1);", "12");
jt!(double_parentheses, "System.out.println(((8 - 6) + 10) / 4);", "3");
