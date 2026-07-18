use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(while_counts_0_to_4, "int i = 0; int sum = 0; while (i < 5) { sum += i; i++; } System.out.println(sum);", "10");
jt!(while_counts_skip_even, "int i = 0; int sum = 0; while (i < 6) { if ((i & 1) == 0) { sum += i; } i++; } System.out.println(sum);", "6");
jt!(while_with_break, "int i = 0; int sum = 0; while (true) { if (i == 4) break; sum += i; i++; } System.out.println(sum);", "6");
jt!(while_with_continue, "int i = 0; int sum = 0; while (i < 6) { i++; if ((i & 1) == 0) continue; sum += i; } System.out.println(sum);", "9");
jt!(nested_while_product, "int i = 0; int sum = 0; while (i < 3) { int j = 0; while (j < 2) { sum += (i + j); j++; } i++; } System.out.println(sum);", "9");
jt!(while_conditionally_assigns, "int i = 0; int x = 0; while (i < 3) { if (i == 1) x = 8; i++; } System.out.println(x);", "8");
jt!(while_false_immediate, "int x = 0; while (false) x++; System.out.println(x);", "0");
jt!(while_true_invariant, "int i = 0; int c = 0; while (i < 1) { c++; i++; } System.out.println(c);", "1");
jt!(while_mutating_step, "int i = 1; int p = 1; while (i < 10) { p *= i; i += 2; } System.out.println(p);", "1");
jt!(while_with_boolean_flag, "int i = 0; int c = 0; boolean done = false; while (!done) { c++; if (c == 3) done = true; } System.out.println(c);", "3");
jt!(while_reassign_limit, "int limit = 2; int i = 0; int c = 0; while (i < limit) { c += 2; if (c > 3) limit = 10; i++; } System.out.println(c);", "6");
jt!(while_nested_sum, "int i = 0; int sum = 0; while (i < 4) { int j = 0; while (j <= i) { sum += j; j++; } i++; } System.out.println(sum);", "20");
jt!(while_updates_string, "int i = 0; String s = \"\"; while (i < 3) { s += \"x\"; i++; } System.out.println(s);", "xxx");
jt!(while_negatives, "int i = -2; int s = 0; while (i < 1) { s += i; i++; } System.out.println(s);", "-3");
jt!(while_or_condition, "int i = 0; int c = 0; while (i < 3 || c < 1) { c += i; i += 2; } System.out.println(c);", "4");
jt!(while_with_side_effect, "int i = 0; int[] arr = {1, 2, 3}; int sum = 0; while (i < arr.length) { if (arr[i] > 1) sum += arr[i]; i++; } System.out.println(sum);", "5");
jt!(while_division, "int i = 1; int c = 0; while (i < 9) { c += (10 / i); i += 3; } System.out.println(c);", "13");
jt!(while_modulo_filter, "int i = 0; int c = 0; while (i < 10) { if (i % 3 == 0) c++; i++; } System.out.println(c);", "4");
jt!(while_composed_condition, "int i = 0; int a = 0; while (i < 6 && a < 3) { i++; if ((i & 1) == 0) a++; } System.out.println(i + \",\" + a);", "6,3");
jt!(while_ternary_condition, "int i = 0; int c = 0; while (i < (c < 3 ? 5 : 0)) { c++; i++; } System.out.println(c);", "3");
jt!(while_compare_types, "int i = 0; long limit = 3; int c = 0; while (i < limit) { c += i; i++; } System.out.println(c);", "3");
jt!(while_double_guard, "int i = 0; double d = 1.5; while (i < 3 && d > 0) { d -= 0.5; i++; } System.out.println((int)d);", "1");
jt!(while_nested_labels, "int x = 0; int y = 0; while (x < 3) { int z = 0; while (z < 2) { y += x + z; z++; } x++; } System.out.println(y);", "9");
jt!(while_array_index, "int[] values = {2, 4, 6}; int i = 0; int s = 1; while (i < values.length) { s *= values[i]; i++; } System.out.println(s);", "48");
jt!(while_empty_body, "int i = 0; while (i < 4) { i++; } System.out.println(i);", "4");
jt!(while_post_increment_in_body, "int i = 0; int s = 0; while (i < 4) { s += i; i++; } System.out.println(s);", "6");
jt!(while_with_return_like_break, "int i = 0; int s = 0; while (i < 10) { s++; if (s == 4) { break; } i++; } System.out.println(s);", "4");

