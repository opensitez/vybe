use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(null_equals_true, "String s = null; System.out.println(s == null);", "true");
jt!(null_not_equals, "String s = \"\"; System.out.println(s != null);", "true");
jt!(null_inequality, "String s = null; System.out.println(s != \"x\");", "true");
jt!(null_with_ternary, "String s = null; System.out.println(s == null ? 1 : 0);", "1");
jt!(null_reference_guard, "String s = null; int n = 0; if (s != null) { n = s.length(); } System.out.println(n);", "0");
jt!(null_reference_else, "String s = null; int n = (s == null) ? 1 : s.length(); System.out.println(n);", "1");
jt!(null_two_refs, "String a = null; String b = null; System.out.println(a == b);", "true");
jt!(null_and_non_null, "String a = null; String b = \"x\"; System.out.println(a != b);", "true");
jt!(null_in_object, "Object o = null; System.out.println(o == null);", "true");
jt!(null_object_check_in_loop, "Object[] arr = {null, new Object(), null}; int c = 0; for (int i = 0; i < arr.length; i++) { if (arr[i] == null) c++; } System.out.println(c);", "2");
jt!(null_else_branch, "Object o = null; String s = (o == null) ? \"none\" : \"some\"; System.out.println(s);", "none");
jt!(null_array, "Object[] a = new Object[]{null, null, new Object()}; int c = 0; for (Object o : a) { if (o == null) c++; } System.out.println(c);", "2");
jt!(null_string_length_safe, "String s = null; int n = (s == null ? 0 : s.length()); System.out.println(n);", "0");
jt!(null_string_or, "String s = null; String t = s == null ? \"x\" : s; System.out.println(t);", "x");
jt!(null_boolean_and, "String s = null; boolean b = s == null && true; System.out.println(b);", "true");
jt!(null_boolean_or, "String s = null; boolean b = s != null || false; System.out.println(b);", "false");
jt!(null_in_switch_like_chain, "String s = null; int v = (s == null) ? 1 : 2; switch (v) { case 1: v = 3; break; default: v = 4; } System.out.println(v);", "3");
jt!(null_then_non_null, "Object o = null; String s = (o == null) ? \"A\" : \"B\"; o = new Object(); s = (o == null) ? \"A\" : \"B\"; System.out.println(s);", "B");
jt!(null_equality_with_cast, "Object o = null; Object p = o; System.out.println(p == o);", "true");
jt!(null_and_instanceof, "Object o = null; System.out.println(o instanceof String);", "false");
jt!(non_null_instanceof, "Object o = \"x\"; System.out.println(o instanceof String);", "true");
jt!(null_string_builder, "String s = null; String t = String.valueOf(s); System.out.println(\"\" + t);", "null");
jt!(null_ternary_string, "String s = null; String t = s == null ? \"yes\" : \"no\"; System.out.println(t);", "yes");
jt!(null_after_assignment, "String s = \"x\"; s = null; System.out.println(s == null ? 1 : 0);", "1");
jt!(null_reference_count, "Object[] items = {null, new Object(), new Object(), null}; int none = 0; int some = 0; for (Object o : items) { if (o == null) { none++; } else { some++; } } System.out.println(none + \",\" + some);", "2,2");
jt!(null_and_addition_prevented, "String s = null; int n = (s == null) ? 1 : 2; n += (s == null ? 3 : 4); System.out.println(n);", "4");
jt!(null_in_if_else, "Object o = null; int v = 0; if (o == null) { v = 9; } else { v = 3; } System.out.println(v);", "9");
jt!(null_chain, "Object[] a = {new Object(), null, new Object()}; int n = 0; for (int i = 0; i < a.length; i++) { if (a[i] == null) n++; } System.out.println(n);", "1");
jt!(null_array_index, "Object[] a = null; if (a == null) { System.out.println(1); } else { System.out.println(0); }", "1");
jt!(null_empty_string, "String s = null; String t = (s == null ? \"\" : s); System.out.println(t.length());", "0");
