use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(empty_int_array_length, "int[] a = new int[0]; System.out.println(a.length);", "0");
jt!(int_array_length_three, "int[] a = {1, 2, 3}; System.out.println(a.length);", "3");
jt!(array_index_zero, "int[] a = {10, 20, 30}; System.out.println(a[0]);", "10");
jt!(array_index_last, "int[] a = {10, 20, 30}; System.out.println(a[2]);", "30");
jt!(array_set_index, "int[] a = {1, 2, 3}; a[1] = 9; System.out.println(a[1]);", "9");
jt!(array_sum_loop, "int[] a = {1,2,3,4}; int s = 0; for(int i = 0; i < a.length; i++) s += a[i]; System.out.println(s);", "10");
jt!(array_sum_for_each_style, "int[] a = {2,2,2}; int s = 0; for(int i : a) s += i; System.out.println(s);", "6");
jt!(array_default_values, "int[] a = new int[3]; System.out.println(a[0] + a[1] + a[2]);", "0");
jt!(byte_array_defaults, "byte[] b = new byte[2]; System.out.println(b[0]);", "0");
jt!(long_array_sum, "long[] l = {1L, 2L, 3L}; long s = 0; for(int i = 0; i < l.length; i++) s += l[i]; System.out.println(s);", "6");
jt!(boolean_array_one, "boolean[] b = {true, false}; System.out.println(b[0]);", "true");
jt!(boolean_array_two, "boolean[] b = {true, false}; System.out.println(b[1]);", "false");
jt!(double_array_arithmetic, "double[] d = {1.5, 2.5}; System.out.println(d[0] + d[1]);", "4.0");
jt!(char_array_length, "char[] c = {'a','b','c'}; System.out.println(c.length);", "3");
jt!(char_array_access, "char[] c = {'a','b'}; System.out.println(c[1]);", "b");
jt!(object_array_length, "String[] s = {\"x\", \"y\"}; System.out.println(s.length);", "2");
jt!(object_array_element, "String[] s = {\"x\", \"y\"}; System.out.println(s[0]);", "x");
jt!(array_update_in_loop, "int[] a = {1, 1, 1}; for(int i = 0; i < a.length; i++) a[i] = a[i] * 2; System.out.println(a[0] + a[1] + a[2]);", "6");
jt!(array_cloned_reference, "int[] a = {1,2}; int[] b = a; b[0] = 9; System.out.println(a[0]);", "9");
jt!(array_copy_like, "int[] a = {5,6,7}; int[] b = new int[a.length]; b[0] = a[0]; b[1] = a[1]; b[2] = a[2]; System.out.println(b[2]);", "7");
jt!(array_contains_loop_sum, "int[] a = {3,1,4}; int c = 0; for(int i = 0; i < a.length; i++) if(a[i] == 1) c++; System.out.println(c);", "1");
jt!(array_with_formula, "int[] a = new int[4]; for(int i = 0; i < a.length; i++) a[i] = i * 2; System.out.println(a[3]);", "6");
jt!(array_negative_values, "int[] a = {-1, -2, -3}; System.out.println(a[1]);", "-2");
jt!(array_minimal_print_only_one, "int[] a = {9,8,7}; System.out.println(a[0] + a[2]);", "16");
jt!(array_ref_length_after_resize_impossible, "int[] a = {1}; int[] b = new int[a.length]; b[0] = a[0]; System.out.println(a.length + b.length);", "2");
