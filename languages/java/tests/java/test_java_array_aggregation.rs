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
    sum_three,
    "int[] a = {1,2,3}; int s = 0; for (int v : a) s += v; System.out.println(s);",
    "6"
);
jt!(
    sum_negative,
    "int[] a = {-1,-2,3}; int s = 0; for (int v : a) s += v; System.out.println(s);",
    "0"
);
jt!(
    product_two,
    "int[] a = {2,3,4}; int p = 1; for (int v : a) p *= v; System.out.println(p);",
    "24"
);
jt!(
    min_three,
    "int[] a = {3,1,2}; int m = a[0]; for (int i = 1; i < a.length; i++) if (a[i] < m) m = a[i]; System.out.println(m);",
    "1"
);
jt!(
    max_three,
    "int[] a = {3,1,2}; int m = a[0]; for (int i = 1; i < a.length; i++) if (a[i] > m) m = a[i]; System.out.println(m);",
    "3"
);
jt!(
    average_floor,
    "int[] a = {1,2,3}; int s = 0; for (int v : a) s += v; System.out.println(s / a.length);",
    "2"
);
jt!(
    count_even,
    "int[] a = {1,2,3,4,5}; int c = 0; for (int v : a) if ((v & 1) == 0) c++; System.out.println(c);",
    "2"
);
jt!(
    count_pos,
    "int[] a = {1,-2,3,-4}; int c = 0; for (int v : a) if (v > 0) c++; System.out.println(c);",
    "2"
);
jt!(
    contains_zero,
    "int[] a = {1,2,0}; int c = 0; for (int v : a) if (v == 0) c++; System.out.println(c);",
    "1"
);
jt!(
    dot_product,
    "int[] a = {1,2,3}; int[] b = {2,1,1}; int p = 0; for (int i = 0; i < a.length; i++) p += a[i] * b[i]; System.out.println(p);",
    "7"
);
jt!(
    prefix_sum_1,
    "int[] a = {1,2,3}; int s = 0; int[] p = new int[a.length]; for (int i = 0; i < a.length; i++) { s += a[i]; p[i] = s; } System.out.println(p[2]);",
    "6"
);
jt!(
    suffix_sum_1,
    "int[] a = {1,2,3,4}; int s = 0; for (int i = a.length - 1; i >= 0; i--) s += a[i]; System.out.println(s);",
    "10"
);
jt!(
    range_sum_positive,
    "int[] a = new int[5]; for (int i = 0; i < a.length; i++) a[i] = i; int s = 0; for (int v : a) s += v; System.out.println(s);",
    "10"
);
jt!(
    range_product_small,
    "int[] a = new int[3]; for (int i = 1; i <= a.length; i++) a[i - 1] = i; int p = 1; for (int v : a) p *= v; System.out.println(p);",
    "6"
);
jt!(
    length_times_two,
    "int[] a = {1,2,3}; System.out.println(a.length * 2);",
    "6"
);
jt!(
    all_greater_than_0,
    "int[] a = {1,2,3}; boolean ok = true; for (int v : a) if (v <= 0) ok = false; System.out.println(ok);",
    "true"
);
jt!(
    any_negative,
    "int[] a = {1,-1,2}; boolean ok = false; for (int v : a) if (v < 0) ok = true; System.out.println(ok);",
    "true"
);
jt!(
    sum_with_index,
    "int[] a = {5,5,5}; int s = 0; for (int i = 0; i < a.length; i++) s += a[i] + i; System.out.println(s);",
    "18"
);
jt!(
    count_distinct_guess,
    "int[] a = {1,1,2,2,3}; int c = 0; for (int i = 0; i < a.length; i++) { boolean seen = false; for (int j = 0; j < i; j++) if (a[j] == a[i]) seen = true; if (!seen) c++; } System.out.println(c);",
    "3"
);
jt!(
    sum_of_odds,
    "int[] a = {1,2,3,4,5}; int s = 0; for (int v : a) if ((v & 1) == 1) s += v; System.out.println(s);",
    "9"
);
jt!(
    sum_of_evens,
    "int[] a = {1,2,3,4,5}; int s = 0; for (int v : a) if ((v & 1) == 0) s += v; System.out.println(s);",
    "6"
);
jt!(
    index_of_max,
    "int[] a = {1,4,2,7,5}; int idx = 0; for (int i = 1; i < a.length; i++) if (a[i] > a[idx]) idx = i; System.out.println(idx);",
    "3"
);
jt!(
    sum_two_arrays,
    "int[] a = {1,2}; int[] b = {3,4}; int s = 0; for (int i = 0; i < a.length; i++) s += a[i] + b[i]; System.out.println(s);",
    "10"
);
jt!(
    running_delta,
    "int[] a = {5,1,1}; int p = a[0]; for (int i = 1; i < a.length; i++) p = p > a[i] ? p : a[i]; System.out.println(p);",
    "5"
);
jt!(
    running_max,
    "int[] a = {1,4,2,8}; int p = a[0]; for (int i = 1; i < a.length; i++) p = p > a[i] ? p : a[i]; System.out.println(p);",
    "8"
);
jt!(
    max_gap,
    "int[] a = {1,9,2,4}; int m = 0; for (int i = 1; i < a.length; i++) m = Math.max(m, Math.abs(a[i] - a[i-1])); System.out.println(m);",
    "8"
);
jt!(
    sum_modulo,
    "int[] a = {10,20,30}; int s = 0; for (int v : a) s += v % 7; System.out.println(s);",
    "6"
);
jt!(
    pairwise_equal,
    "int[] a = {1,2,3,2,1}; boolean ok = true; for (int i = 0; i < a.length/2; i++) if (a[i] != a[a.length -1 -i]) ok = false; System.out.println(ok);",
    "false"
);
jt!(
    pairwise_palindrome,
    "int[] a = {1,2,2,1}; boolean ok = true; for (int i = 0; i < a.length/2; i++) if (a[i] != a[a.length -1 -i]) ok = false; System.out.println(ok);",
    "true"
);
jt!(
    sum_abs,
    "int[] a = {-1,-2,3}; int s = 0; for (int v : a) s += Math.abs(v); System.out.println(s);",
    "6"
);
