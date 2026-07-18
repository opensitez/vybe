use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(plus_equals, "int a = 1; a += 2; System.out.println(a);", "3");
jt!(minus_equals, "int a = 10; a -= 4; System.out.println(a);", "6");
jt!(times_equals, "int a = 3; a *= 4; System.out.println(a);", "12");
jt!(divide_equals, "int a = 20; a /= 4; System.out.println(a);", "5");
jt!(mod_equals, "int a = 17; a %= 5; System.out.println(a);", "2");
jt!(and_equals, "int a = 6; a &= 3; System.out.println(a);", "2");
jt!(or_equals, "int a = 1; a |= 4; System.out.println(a);", "5");
jt!(xor_equals, "int a = 7; a ^= 3; System.out.println(a);", "4");
jt!(left_shift_equals, "int a = 3; a <<= 2; System.out.println(a);", "12");
jt!(right_shift_equals, "int a = 16; a >>= 2; System.out.println(a);", "4");
jt!(unsigned_shift_equals, "int a = -1; a >>>= 30; System.out.println(a);", "3");
jt!(for_loop_with_plus, "int a = 0; for (int i = 0; i < 3; i++) { a += i; } System.out.println(a);", "3");
jt!(array_plus_assign, "int[] values = {1,2,3}; values[0] += values[1]; values[1] *= 2; System.out.println(values[0] + \":\" + values[1]);", "3:4");
jt!(compare_after_update, "int a = 5; a *= 2; boolean ok = (a == 10); System.out.println(ok);", "true");
jt!(mixed_operations, "int a = 1; a += 2; a *= 3; a -= 1; a /= 2; System.out.println(a);", "4");
jt!(nested_assignment_target, "int a = 1; a += (a += 1); System.out.println(a);", "3");
jt!(bit_mix, "int a = 12; a ^= 3; a &= 10; a |= 1; System.out.println(a);", "11");
jt!(assign_array_in_loop, "int[] arr = {1, 2, 3}; int sum = 0; for (int i = 0; i < arr.length; i++) { sum += arr[i]; arr[i] *= 2; } System.out.println(sum + \":\" + arr[2]);", "6:6");
jt!(assign_in_branch_true, "int a = 0; if (true) { a += 5; } else { a += 2; } System.out.println(a);", "5");
jt!(assign_in_branch_false, "int a = 0; if (false) { a += 5; } else { a += 2; } System.out.println(a);", "2");
jt!(assign_with_method_call, "int a = 1; a += Integer.parseInt(\"2\"); System.out.println(a);", "3");
jt!(assign_three_terms, "int a = 10; a /= 2; a %= 2; a += 4; System.out.println(a);", "1");
jt!(shift_combo, "int a = 8; a >>>= 1; a <<= 1; System.out.println(a);", "8");
jt!(while_update_chain, "int a = 0; int i = 0; while (i < 4) { a += i; a += 1; i++; } System.out.println(a);", "10");
jt!(plus_then_negate, "int a = 1; a += 2; a = -a; System.out.println(a);", "-3");
jt!(divide_before_store, "int a = 20; int b = 4; b = (a /= 2); System.out.println(b + \":\" + a);", "10:10");
