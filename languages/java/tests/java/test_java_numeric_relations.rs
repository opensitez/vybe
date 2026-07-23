use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(equal_true, "System.out.println(3 == 3);", "true");
jt!(equal_false, "System.out.println(3 == 4);", "false");
jt!(not_equal_true, "System.out.println(3 != 4);", "true");
jt!(not_equal_false, "System.out.println(3 != 3);", "false");
jt!(less_true, "System.out.println(3 < 4);", "true");
jt!(less_false, "System.out.println(4 < 3);", "false");
jt!(greater_true, "System.out.println(4 > 3);", "true");
jt!(greater_false, "System.out.println(3 > 4);", "false");
jt!(le_true, "System.out.println(3 <= 3);", "true");
jt!(le_false, "System.out.println(4 <= 3);", "false");
jt!(ge_true, "System.out.println(3 >= 3);", "true");
jt!(ge_false, "System.out.println(2 >= 3);", "false");
jt!(
    compare_ints,
    "int a = 2; int b = 3; System.out.println(a < b);",
    "true"
);
jt!(compare_char, "System.out.println('a' < 'b');", "true");
jt!(
    compare_string_length,
    "String a = \"abc\"; String b = \"abcd\"; System.out.println(a.length() < b.length());",
    "true"
);
jt!(
    compare_object_ref_same,
    "Object a = new Object(); Object b = a; System.out.println(a == b);",
    "true"
);
jt!(
    compare_object_ref_diff,
    "Object a = new Object(); Object b = new Object(); System.out.println(a == b);",
    "false"
);
jt!(
    compare_null_ref,
    "Object a = null; System.out.println(a == null);",
    "true"
);
jt!(
    not_compare_null_ref,
    "Object a = null; System.out.println(a != null);",
    "false"
);
jt!(
    float_compare_eq,
    "System.out.println(1.0f == 1.0f);",
    "true"
);
jt!(
    double_compare_near,
    "System.out.println(0.1 + 0.2 == 0.3);",
    "false"
);
jt!(
    ternary_relation,
    "System.out.println(2 + 2 == 4 ? \"yes\" : \"no\");",
    "yes"
);
jt!(
    chain_relation_1,
    "System.out.println(1 < 2 && 2 < 3);",
    "true"
);
jt!(
    chain_relation_2,
    "System.out.println(1 > 2 || 2 < 3);",
    "true"
);
jt!(
    chain_relation_3,
    "System.out.println((1 == 1) && (1 != 2));",
    "true"
);
jt!(
    relation_in_loop,
    "int c = 0; for (int i = 0; i < 5; i++) { if (i >= 2) c++; } System.out.println(c);",
    "3"
);
jt!(
    relation_with_assignment,
    "int x = 1; boolean b = (x += 1) == 2; System.out.println(b);",
    "true"
);
jt!(
    relation_with_mod,
    "int x = 5; boolean b = x % 2 == 1; System.out.println(b);",
    "true"
);
jt!(
    unsigned_compare_sim,
    "System.out.println(Integer.compareUnsigned(-1, 1) > 0);",
    "true"
);
jt!(
    long_relation,
    "long a = 10000000000L; long b = 10000000000L; System.out.println(a >= b);",
    "true"
);
