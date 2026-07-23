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
    fill_array_with_index,
    "int[] a = new int[3]; for (int i = 0; i < a.length; i++) { a[i] = i; } System.out.println(a[2]);",
    "2"
);
jt!(
    sum_after_fill,
    "int[] a = new int[4]; for (int i = 0; i < a.length; i++) { a[i] = i + 1; } int s = 0; for (int i = 0; i < a.length; i++) { s += a[i]; } System.out.println(s);",
    "10"
);
jt!(
    copy_into_new_array,
    r#"int[] a = {1,2,3}; int[] b = {0,0,0}; for (int i = 0; i < a.length; i++) { b[i] = a[i]; } System.out.println(b[0] + "," + b[2]);"#,
    "1,3"
);
jt!(
    reverse_in_place,
    r#"int[] a = {1,2,3,4}; for (int i = 0; i < a.length / 2; i++) { int t = a[i]; a[i] = a[a.length - 1 - i]; a[a.length - 1 - i] = t; } System.out.println(a[0] + "," + a[3]);"#,
    "4,1"
);
jt!(
    reverse_sum_after,
    "int[] a = {1,2,3}; int[] b = new int[a.length]; for (int i = 0; i < a.length; i++) { b[i] = a[a.length - 1 - i]; } int s=0; for (int v : b) s += v; System.out.println(s);",
    "6"
);
jt!(
    find_max,
    "int[] a = {4,9,1,6}; int max = a[0]; for (int i = 1; i < a.length; i++) { if (a[i] > max) max = a[i]; } System.out.println(max);",
    "9"
);
jt!(
    find_min,
    "int[] a = {4,9,1,6}; int min = a[0]; for (int i = 1; i < a.length; i++) { if (a[i] < min) min = a[i]; } System.out.println(min);",
    "1"
);
jt!(
    count_matches,
    "int[] a = {1,2,2,3,2}; int c=0; for (int i = 0; i < a.length; i++) { if (a[i] == 2) c++; } System.out.println(c);",
    "3"
);
jt!(
    index_of_value,
    "int[] a = {5,6,7}; int idx = -1; for (int i = 0; i < a.length; i++) { if (a[i] == 7) idx = i; } System.out.println(idx);",
    "2"
);
jt!(
    shift_left_by_one,
    r#"int[] a = {1,2,3}; int first = a[0]; for (int i = 0; i < a.length - 1; i++) { a[i] = a[i + 1]; } a[a.length - 1] = first; System.out.println(a[0] + "," + a[2]);"#,
    "2,1"
);
jt!(
    sum_with_filter,
    "int[] a = {1,2,3,4}; int s=0; for (int v : a) { if (v % 2 == 0) s += v; } System.out.println(s);",
    "6"
);
jt!(
    pairwise_sum_array,
    "int[] a = {1,2,3,4}; int[] b = {2,2,2,2}; int c = 0; for (int i = 0; i < a.length; i++) { c += a[i] + b[i]; } System.out.println(c);",
    "18"
);
jt!(
    rotate_right_one,
    r#"int[] a = {1,2,3}; int last = a[a.length - 1]; for (int i = a.length - 1; i > 0; i--) a[i] = a[i - 1]; a[0] = last; System.out.println(a[0] + "," + a[1] + "," + a[2]);"#,
    "3,1,2"
);
jt!(
    prefix_increment_array,
    "int[] a = {1,2,3}; for (int i = 0; i < a.length; i++) { a[i] = a[i] + 1; } System.out.println(a[2]);",
    "4"
);
jt!(
    suffix_increment_array,
    "int[] a = {1,2,3}; for (int i = 0; i < a.length; i++) { a[i]++; } System.out.println(a[1]);",
    "3"
);
jt!(
    multiply_by_index,
    "int[] a = {1,1,1,1}; for (int i = 0; i < a.length; i++) { a[i] = a[i] * i; } System.out.println(a[3]);",
    "3"
);
jt!(
    clone_like_sum,
    "int[] a = {2,4,6}; int[] b = {0,0,0}; for (int i = 0; i < a.length; i++) b[i] = a[i]; int s = 0; for (int v : b) s += v; System.out.println(s);",
    "12"
);
jt!(
    prefix_scan,
    "int[] a = {1,2,3}; int[] p = {0,0,0}; int s = 0; for (int i = 0; i < a.length; i++) { s += a[i]; p[i] = s; } System.out.println(p[2]);",
    "6"
);
jt!(
    all_equal_to_value,
    "int[] a = {2,2,2}; boolean ok = true; for (int v : a) { if (v != 2) ok = false; } System.out.println(ok);",
    "true"
);
jt!(
    first_and_last,
    r#"int[] a = {9,8,7}; System.out.println(a[0] + "," + a[a.length - 1]);"#,
    "9,7"
);
jt!(
    sum_of_squares,
    "int[] a = {1,2,3}; int s = 0; for (int v : a) { s += v * v; } System.out.println(s);",
    "14"
);
jt!(
    count_longer_than_first,
    "int[] a = {5,7,3,9}; int c = 0; int first = a[0]; for (int i = 1; i < a.length; i++) { if (a[i] > first) c++; } System.out.println(c);",
    "2"
);
jt!(
    zero_fill,
    "int[] a = {1,2,3}; for (int i = 0; i < a.length; i++) a[i] = 0; System.out.println(a[1]);",
    "0"
);
jt!(
    merge_last_two,
    r#"int[] a = {1,2,3}; int[] b = {4,5,6}; int[] c = {a[1], b[1]}; System.out.println(c[0] + "," + c[1]);"#,
    "2,5"
);
jt!(
    accumulate_even_then_odd,
    r#"int[] a = {1,2,3,4}; int even=0, odd=0; for (int v : a) { if (v % 2 == 0) even += v; else odd += v; } System.out.println(even + "," + odd);"#,
    "6,4"
);
