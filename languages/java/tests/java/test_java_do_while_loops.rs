use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(executes_once, "int count = 0; do { count++; } while (false); System.out.println(count);", "1");
jt!(executes_three_times, "int count = 0; int i = 0; do { count += 1; i++; } while (i < 3); System.out.println(count);", "3");
jt!(increments_sum, "int sum = 0; int i = 0; do { sum += i; i++; } while (i < 4); System.out.println(sum);", "6");
jt!(break_stops_loop, "int sum = 0; int i = 0; do { sum += i; i++; if (i == 2) break; } while (i < 10); System.out.println(sum);", "1");
jt!(continue_skips, "int sum = 0; int i = 0; do { i++; if (i == 2) continue; sum += i; } while (i < 4); System.out.println(sum);", "7");
jt!(nested_counter, "int outer = 0; int i = 0; do { int j = 0; outer += i; i++; } while (i < 3); System.out.println(outer);", "3");
jt!(boolean_exit, "int i = 0; int sum = 0; do { sum += 2; i++; } while ((i < 2) && (sum < 10)); System.out.println(sum);", "4");
jt!(do_while_with_if, "int x = 0; int i = 0; do { if (i % 2 == 0) { x++; } i++; } while (i < 5); System.out.println(x);", "3");
jt!(do_while_mutating_condition, "int i = 0; int x = 0; do { x += i; i *= 2; i++; } while (i < 4); System.out.println(x);", "0");
jt!(long_term, "int total = 0; int i = 1; do { total += i; i *= 2; } while (i <= 8); System.out.println(total);", "15");
jt!(negative_start, "int i = -1; int total = 0; do { total += i; i++; } while (i < 2); System.out.println(total);", "0");
jt!(string_counter, "int i = 0; String s = \"\"; do { s += i; i++; } while (i < 3); System.out.println(s);", "012");
jt!(flag_guard, "int i = 0; int x = 0; boolean done = false; do { x++; if (x == 2) done = true; } while (!done && (i++ < 3)); System.out.println(x);", "2");
jt!(double_stage, "int x = 0; int i = 0; do { x += i; i++; } while (i < 2); do { x += i; i++; } while (i < 4); System.out.println(x);", "5");
jt!(nested_do_while, "int total = 0; int i = 0; do { int j = 0; do { total += j; j++; } while (j < 2); i++; } while (i < 3); System.out.println(total);", "6");
jt!(do_while_with_array, "int[] a = {1, 2, 3}; int i = 0; int sum = 0; do { sum += a[i]; i++; } while (i < a.length); System.out.println(sum);", "6");
jt!(do_while_deep_assignments, "int a = 1; int b = 0; do { b = a * 2; a++; } while (a < 4); System.out.println(a + \",\" + b);", "4,6");
jt!(break_nested, "int a = 0; int b = 0; do { a++; if (a == 1) { b++; continue; } if (a == 3) { break; } b += a; } while (a < 5); System.out.println(a + \",\" + b);", "3,1");
jt!(while_like_count, "int i = 0; int x = 0; do { x += 1; i++; } while (i < 5); System.out.println(x);", "5");
jt!(do_while_condition_false_late, "int i = 0; int x = 0; do { x += i; i++; } while (x < 0); System.out.println(i);", "1");
jt!(conditional_in_condition, "int i = 0; int x = 0; do { x += i; i++; } while ((i < 2) ? true : false); System.out.println(x);", "1");
jt!(modulo_in_body, "int i = 0; int c = 0; do { c += i % 3; i++; } while (i < 5); System.out.println(c);", "10");
jt!(assignment_condition, "int i = 0; int x = 0; do { x++; } while ((i = 1) == 0); System.out.println(i);", "1");
jt!(char_progression, "int i = 0; char c = 'a'; do { c++; i++; } while (i < 3); System.out.println(c);", "d");
jt!(final_check, "int i = 0; int x = 0; do { x += i; i++; } while (x < 3); System.out.println(i);", "3");

