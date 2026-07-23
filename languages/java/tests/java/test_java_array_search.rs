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
    first_match,
    "int[] a = {1,2,3}; int p = -1; for (int i = 0; i < a.length; i++) { if (a[i] == 2) { p = i; break; } } System.out.println(p);",
    "1"
);
jt!(
    no_match,
    "int[] a = {1,2,3}; int p = -1; for (int i = 0; i < a.length; i++) { if (a[i] == 9) { p = i; break; } } System.out.println(p);",
    "-1"
);
jt!(
    last_match,
    "int[] a = {1,2,3,2}; int p = -1; for (int i = 0; i < a.length; i++) { if (a[i] == 2) p = i; } System.out.println(p);",
    "3"
);
jt!(
    first_gt,
    "int[] a = {1,2,3,4}; int p = -1; for (int i = 0; i < a.length; i++) { if (a[i] > 2) { p = a[i]; break; } } System.out.println(p);",
    "3"
);
jt!(
    count_gt,
    "int[] a = {1,2,3,4}; int c = 0; for (int v : a) if (v > 2) c++; System.out.println(c);",
    "2"
);
jt!(
    count_even,
    "int[] a = {1,2,3,4,5,6}; int c = 0; for (int v : a) if ((v & 1) == 0) c++; System.out.println(c);",
    "3"
);
jt!(
    find_string_length,
    "String[] a = {\"a\",\"bb\",\"ccc\"}; int p = 0; for (int i = 0; i < a.length; i++) { if (a[i].length() == 3) { p = i; break; } } System.out.println(p);",
    "2"
);
jt!(
    find_boolean_true,
    "boolean[] a = {false, false, true}; int p = 0; for (int i = 0; i < a.length; i++) if (a[i]) { p = i; break; } System.out.println(p);",
    "2"
);
jt!(
    all_match,
    "int[] a = {2,2,2}; boolean ok = true; for (int v : a) if (v != 2) ok = false; System.out.println(ok);",
    "true"
);
jt!(
    none_match,
    "int[] a = {1,3,5}; boolean ok = true; for (int v : a) if (v % 2 == 0) ok = false; System.out.println(ok);",
    "true"
);
jt!(
    contains_substring_array,
    "String[] s = {\"ab\", \"bc\"}; String t = \"bc\"; boolean ok = false; for (int i = 0; i < s.length; i++) { if (s[i].equals(t)) ok = true; } System.out.println(ok);",
    "true"
);
jt!(
    first_negative,
    "int[] a = {1,-2,3}; int n = 0; for (int i = 0; i < a.length; i++) { if (a[i] < 0) { n = a[i]; break; } } System.out.println(n);",
    "-2"
);
jt!(
    min_at_or_above,
    "int[] a = {3,1,2}; int m = a[0]; for (int v : a) if (v < m) m = v; System.out.println(m);",
    "1"
);
jt!(
    all_between,
    "int[] a = {1,2,3}; boolean ok = true; for (int v : a) if (v < 0 || v > 5) ok = false; System.out.println(ok);",
    "true"
);
jt!(
    count_strings_with_a,
    "String[] s = {\"cat\",\"dog\",\"bat\"}; int c = 0; for (int i = 0; i < s.length; i++) if (s[i].contains(\"a\")) c++; System.out.println(c);",
    "2"
);
jt!(
    sum_indexed,
    "int[] a = {1,2,3}; int s = 0; for (int i = 0; i < a.length; i++) if (i > 0) s += a[i]; System.out.println(s);",
    "5"
);
jt!(
    first_ge_3_index,
    "int[] a = {1,2,4,5}; int i = 0; while (i < a.length && a[i] < 3) i++; System.out.println(i);",
    "2"
);
jt!(
    search_char_in_strings,
    "String[] s = {\"aa\",\"bb\",\"ca\"}; int n = 0; for (int i = 0; i < s.length; i++) if (s[i].contains(\"c\")) n++; System.out.println(n);",
    "1"
);
jt!(
    binary_search_small,
    "int[] a = {1,2,3,4,5}; int t = 3; int l = 0; int r = a.length - 1; int p = -1; while (l <= r) { int m = (l + r) / 2; if (a[m] == t) { p = m; break; } if (a[m] < t) l = m + 1; else r = m - 1; } System.out.println(p);",
    "2"
);
jt!(
    find_first_nonzero,
    "int[] a = {0,0,3,0}; int p = 0; for (int i = 0; i < a.length; i++) { if (a[i] != 0) { p = i; break; } } System.out.println(p);",
    "2"
);
jt!(
    count_unique_adjacent,
    "int[] a = {1,1,2,2,3,3,3}; int c = 0; for (int i = 0; i < a.length; i++) { if (i == 0 || a[i] != a[i - 1]) c++; } System.out.println(c);",
    "4"
);
jt!(
    search_from_end,
    "int[] a = {1,2,3,2,1}; int p = -1; for (int i = a.length - 1; i >= 0; i--) { if (a[i] == 2) { p = i; break; } } System.out.println(p);",
    "3"
);
jt!(
    contains_all_small,
    "int[] a = {1,2,3}; boolean ok1 = false, ok2 = false, ok3 = false; for (int v : a) { if (v == 1) ok1 = true; if (v == 2) ok2 = true; if (v == 3) ok3 = true; } System.out.println((ok1 && ok2) && ok3);",
    "true"
);
jt!(
    sum_until_match,
    "int[] a = {1,2,3,4}; int s = 0; for (int i = 0; i < a.length; i++) { if (a[i] == 4) break; s += a[i]; } System.out.println(s);",
    "6"
);
jt!(
    search_minimum_positive,
    "int[] a = {-1,-2,3,4}; int m = 0; boolean found = false; for (int v : a) { if (v > 0) { m = v; found = true; break; } } System.out.println(found ? m : 0);",
    "3"
);
jt!(
    has_exact_two,
    "int[] a = {1,2,2,3}; int c = 0; for (int v : a) if (v == 2) c++; System.out.println(c == 2);",
    "true"
);
jt!(
    search_zero_sum_prefix,
    "int[] a = {1,-1,2,-2,0}; int n = 0; for (int i = 0; i < a.length; i++) { n += a[i]; if (n == 0) { System.out.println(i); return; } } System.out.println(-1);",
    "1"
);
