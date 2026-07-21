use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(int_div_trunc_pos, "System.out.println(7 / 2);", "3");
jt!(int_div_trunc_neg, "System.out.println(-7 / 2);", "-3");
jt!(int_div_even, "System.out.println(8 / 2);", "4");
jt!(mod_pos, "System.out.println(7 % 3);", "1");
jt!(mod_neg_num, "System.out.println(-7 % 3);", "-1");
jt!(mod_neg_den, "System.out.println(7 % -3);", "1");
jt!(double_div, "System.out.println((int)(7.0 / 2.0));", "3");
jt!(double_div_precise, "System.out.println(7.0 / 2.0);", "3.5");
jt!(int_div_by_var, "int a = 1; int b = 2; System.out.println(4 / b + a);", "3");
jt!(division_chain, "int a = 10; int b = 2; int c = 3; System.out.println(a / b / c);", "1");
jt!(mod_chain, "int a = 10; int b = 2; int c = 3; System.out.println(a % b % c);", "0");
jt!(division_mixed, "int a = 10; double b = 4.0; System.out.println(a / b);", "2.5");
jt!(mod_mixed, "int a = 10; double b = 4.0; System.out.println(a % b);", "2.0");
jt!(percentile_like, "int[] a = {10,20,30,40}; int s = 0; for (int v : a) { s += v / 10; } System.out.println(s);", "10");
jt!(floor_div_alias, "System.out.println(Math.floorDiv(7, 2));", "3");
jt!(floor_div_neg_alias, "System.out.println(Math.floorDiv(-7, 2));", "-4");
jt!(floor_mod_alias, "System.out.println(Math.floorMod(-7, 2));", "1");
jt!(rational_sum, "System.out.println((1.0 / 3) + (1.0 / 3));", "0.6666666666666666");
jt!(division_zero_handled, "System.out.println(5 / 1);", "5");
jt!(double_div_by_zero, "System.out.println(Double.isInfinite(5.0 / 0.0));", "true");
jt!(mod_zero_mod, "System.out.println(1 % 3);", "1");
jt!(double_mod_zero, "System.out.println(5.5 % 2.0);", "1.5");
jt!(divide_then_add, "System.out.println(9 / 3 + 1);", "4");
jt!(divide_then_multiply, "System.out.println((9 / 3) * 2);", "6");
jt!(divide_precedence, "System.out.println(18 / 3 + 2 * 2);", "10");
jt!(division_precedence_parenthesis, "System.out.println(18 / (3 + 2) * 2);", "6");
jt!(cast_division_before, "System.out.println((int) (10 / 3.0));", "3");
jt!(long_division, "System.out.println(10L / 3L);", "3");
jt!(int_long_mix, "System.out.println(10L / 3);", "3");
jt!(double_cast_rounds, "System.out.println((int) (10 / 4.0));", "2");
