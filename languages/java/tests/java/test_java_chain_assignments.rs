use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(plus_chain, "int a = 1; a += 2; a += 3; System.out.println(a);", "6");
jt!(minus_chain, "int a = 20; a -= 3; a -= 4; System.out.println(a);", "13");
jt!(times_chain, "int a = 2; a *= 3; a *= 2; System.out.println(a);", "12");
jt!(divide_chain, "int a = 80; a /= 2; a /= 2; System.out.println(a);", "20");
jt!(mod_chain, "int a = 17; a %= 10; a %= 3; System.out.println(a);", "1");
jt!(bit_and_chain, "int a = 14; a &= 7; a &= 3; System.out.println(a);", "2");
jt!(bit_or_chain, "int a = 1; a |= 2; a |= 4; System.out.println(a);", "7");
jt!(bit_xor_chain, "int a = 15; a ^= 3; a ^= 3; System.out.println(a);", "15");
jt!(left_shift_chain, "int a = 1; a <<= 2; a <<= 1; System.out.println(a);", "8");
jt!(right_shift_chain, "int a = 32; a >>= 2; a >>= 1; System.out.println(a);", "4");
jt!(unsigned_right_shift_chain, "int a = -16; a >>>= 1; a >>>= 1; System.out.println(a > 0);", "true");
jt!(compound_plus_and_copy, "int a = 1; int b = a; b += 4; a += b; System.out.println(a);", "6");
jt!(assignment_in_conditional_true, "int a = 1; boolean flag = true; a = flag ? a += 2 : 3; System.out.println(a);", "3");
jt!(assignment_in_conditional_false, "int a = 1; boolean flag = false; a = flag ? 3 : (a += 2); System.out.println(a);", "3");
jt!(array_index_assignment, "int[] a = {1,2,3}; a[1] += 4; System.out.println(a[1]);", "6");
jt!(reference_update_chain, "int[] box = {1}; box[0] += 2; box[0] += 3; System.out.println(box[0]);", "6");
jt!(assignment_on_decl_init, "int a = 1; int b = a; b += 2; System.out.println(b);", "3");
jt!(nested_for_assignment, "int a = 0; for (int i = 0; i < 3; i++) { a += i; a += 1; } System.out.println(a);", "6");
jt!(while_assignment, "int a = 0; int i = 1; while (i <= 3) { a += i; a += 2; i++; } System.out.println(a);", "9");
jt!(from_expression, "int a = 0; a += a + 2 + 1; System.out.println(a);", "3");
jt!(pre_post_ops_noted, "int a = 1; ++a; a += 3; System.out.println(a);", "5");
jt!(post_then_assign, "int a = 0; a += 1; a += a++; System.out.println(a);", "2");
jt!(assignment_with_brace_scope, "int a = 1; { int b = a; a = b; } a += 4; System.out.println(a);", "5");
jt!(mixed_primitive_assignments, "int a = 2; a += 3; a -= 1; a *= 2; a /= 2; a %= 3; System.out.println(a);", "3");
jt!(multiple_composite_updates, "int a = 10; a = (a += 1) + (a += 2); System.out.println(a);", "24");
jt!(assignment_from_ternary_object, "int a = 2; int b = (a += 1) > 2 ? (a += 2) : (a += 3); System.out.println(a + ";" + b);", "5;5");
jt!(bitwise_assign_after_branch, "int a = 8; boolean ok = true; if (ok) { a &= 6; } else { a |= 6; } System.out.println(a);", "0");
