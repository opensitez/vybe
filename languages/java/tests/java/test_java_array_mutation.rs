use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(set_first, "int[] a = {1,2,3}; a[0] = 9; System.out.println(a[0]);", "9");
jt!(set_last, "int[] a = {1,2,3}; a[2] = 7; System.out.println(a[2]);", "7");
jt!(increment_each, "int[] a = {1,1,1}; for (int i = 0; i < a.length; i++) a[i] += 1; System.out.println(a[0] + a[1] + a[2]);", "6");
jt!(swap_first_last, "int[] a = {1,2,3}; int t = a[0]; a[0] = a[2]; a[2] = t; System.out.println(a[0] + \",\" + a[2]);", "3,1");
jt!(fill_with_index, "int[] a = new int[4]; for (int i = 0; i < a.length; i++) a[i] = i * 2; System.out.println(a[3]);", "6");
jt!(reverse_three, "int[] a = {1,2,3}; for (int i = 0; i < a.length / 2; i++) { int t = a[i]; a[i] = a[a.length -1 -i]; a[a.length -1 -i] = t; } System.out.println(a[0] + a[2]);", "4");
jt!(left_shift, "int[] a = {1,2,3,4}; for (int i = 0; i < a.length - 1; i++) a[i] = a[i + 1]; a[a.length -1] = 0; System.out.println(a[0] + a[3]);", "2");
jt!(right_shift, "int[] a = {1,2,3,4}; for (int i = a.length -1; i > 0; i--) a[i] = a[i - 1]; a[0] = 0; System.out.println(a[0] + a[1]);", "2");
jt!(double_values, "int[] a = {1,2,3}; for (int i = 0; i < a.length; i++) a[i] = a[i] * 2; System.out.println(a[2]);", "6");
jt!(halve_odd, "int[] a = {1,3,5}; for (int i = 0; i < a.length; i++) a[i] = a[i] / 2; System.out.println(a[1]);", "1");
jt!(filter_to_zero, "int[] a = {1,2,3,4}; for (int i = 0; i < a.length; i++) if (a[i] % 2 == 1) a[i] = 0; System.out.println(a[1] + \",\" + a[2]);", "2,0");
jt!(count_nonzero, "int[] a = {0,2,0,4}; int c = 0; for (int i = 0; i < a.length; i++) { if (a[i] != 0) c++; a[i] = 9; } System.out.println(c);", "2");
jt!(matrix_mutate_row, "int[][] a = {{1,2},{3,4}}; a[0][0] = a[1][1]; System.out.println(a[0][0]);", "4");
jt!(rotate_values, "int[] a = {1,2,3,4,5}; int t = a[0]; for (int i = 0; i < a.length -1; i++) a[i] = a[i + 1]; a[a.length -1] = t; System.out.println(a[0] + \",\" + a[4]);", "2,1");
jt!(set_range_small, "int[] a = {0,0,0,0}; for (int i = 1; i < 3; i++) a[i] = i; System.out.println(a[1] + a[2]);", "3");
jt!(copy_reference, "int[] a = {1,2,3}; int[] b = a; b[1] = 9; System.out.println(a[1]);", "9");
jt!(copy_after_clone, "int[] a = {1,2,3}; int[] b = a.clone(); b[1] = 9; System.out.println(a[1] + \",\" + b[1]);", "2,9");
jt!(set_even, "int[] a = {1,2,3,4}; for (int i = 0; i < a.length; i++) if ((i & 1) == 0) a[i] = 0; System.out.println(a[0] + a[1] + a[2] + a[3]);", "8");
jt!(scale_if_gt2, "int[] a = {1,2,3,4}; for (int i = 0; i < a.length; i++) if (a[i] > 2) a[i] *= 2; System.out.println(a[2] + \",\" + a[3]);", "6,8");
jt!(sum_in_place, "int[] a = {1,2,3}; for (int i = 1; i < a.length; i++) a[i] += a[i -1]; System.out.println(a[2]);", "6");
jt!(prefix_fill, "int[] a = {1,2,3,4}; int running = 0; for (int i = 0; i < a.length; i++) { running += a[i]; a[i] = running; } System.out.println(a[3]);", "10");
jt!(multiply_mirror, "int[] a = {2,3,4,5}; for (int i = 0; i < a.length / 2; i++) { int j = a.length -1 - i; a[i] = a[i] * a[j]; } System.out.println(a[0] + \",\" + a[1]);", "10,12");
jt!(set_if_null_not, "Object[] a = new Object[3]; for (int i = 0; i < a.length; i++) if (a[i] == null) a[i] = new Object(); System.out.println(a.length);", "3");
jt!(shift_values_down, "int[] a = {1,2,3,4}; for (int i = 1; i < a.length; i++) a[i - 1] = a[i]; System.out.println(a[0] + \",\" + a[2]);", "2,4");
jt!(shift_values_with_zero, "int[] a = {1,2,3,4}; for (int i = a.length -1; i > 0; i--) a[i] = a[i -1]; a[0] = 0; System.out.println(a[0] + \",\" + a[1]);", "0,1");
jt!(zero_tail, "int[] a = {1,2,3}; a[a.length -1] = 0; System.out.println(a[2]);", "0");
jt!(zero_all, "int[] a = {1,2,3}; for (int i = 0; i < a.length; i++) a[i] = 0; System.out.println(a[0] + a[1] + a[2]);", "0");
jt!(increment_prefix, "int[] a = {5,6,7}; for (int i = 0; i < a.length; i++) a[i] += i + 1; System.out.println(a[2]);", "10");

