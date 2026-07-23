use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(eq_true, "System.out.println(5 == 5);", "true");
jt!(eq_false, "System.out.println(5 == 6);", "false");
jt!(neq_true, "System.out.println(5 != 6);", "true");
jt!(neq_false, "System.out.println(6 != 6);", "false");
jt!(lt_true, "System.out.println(3 < 4);", "true");
jt!(lt_false, "System.out.println(4 < 3);", "false");
jt!(lte_true, "System.out.println(3 <= 3);", "true");
jt!(lte_false, "System.out.println(4 <= 3);", "false");
jt!(gt_true, "System.out.println(10 > 2);", "true");
jt!(gt_false, "System.out.println(2 > 10);", "false");
jt!(gte_true, "System.out.println(10 >= 10);", "true");
jt!(gte_false, "System.out.println(9 >= 10);", "false");
jt!(
    string_eq_content,
    "System.out.println(\"ab\".equals(\"ab\"));",
    "true"
);
jt!(
    string_eq_content_false,
    "System.out.println(\"ab\".equals(\"ba\"));",
    "false"
);
jt!(
    reference_equals_same,
    "String a = \"x\"; String b = a; System.out.println(a == b);",
    "true"
);
jt!(
    reference_equals_distinct,
    "String a = new String(\"x\"); String b = new String(\"x\"); System.out.println(a == b);",
    "false"
);
jt!(
    compare_operator_chain,
    "System.out.println(1 < 2 && 2 < 3);",
    "true"
);
jt!(
    compare_operator_break,
    "System.out.println(1 < 2 && 3 < 2);",
    "false"
);
jt!(
    comparison_with_arithmetic,
    "System.out.println((2 + 3) == 5);",
    "true"
);
jt!(
    chained_notation_false,
    "System.out.println(3 > 2 == false);",
    "false"
);
jt!(
    chained_notation_true,
    "System.out.println(3 > 2 == true);",
    "true"
);
jt!(
    mixed_boolean_eq,
    "System.out.println((5 > 2) == true);",
    "true"
);
jt!(
    boolean_casted_from_relational,
    "System.out.println((1 + 1 == 2));",
    "true"
);
jt!(
    string_length_compare,
    "System.out.println(\"abc\".length() == 3);",
    "true"
);
jt!(
    object_eq_wrapper,
    "Integer x = 5; Integer y = 5; System.out.println(x == y);",
    "true"
);
jt!(
    object_not_eq_wrapper,
    "Integer x = 1000; Integer y = 1000; System.out.println(x == y);",
    "false"
);
