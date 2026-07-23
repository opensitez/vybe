use crate::helpers::run_main;

macro_rules! jt {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_main($src), vec![$expected]);
        }
    };
}

jt!(literal_true, "System.out.println(true);", "true");
jt!(literal_false, "System.out.println(false);", "false");
jt!(
    logical_and_true,
    "System.out.println(true && true);",
    "true"
);
jt!(
    logical_and_false,
    "System.out.println(true && false);",
    "false"
);
jt!(
    logical_or_true_left,
    "System.out.println(true || false);",
    "true"
);
jt!(
    logical_or_false,
    "System.out.println(false || false);",
    "false"
);
jt!(logical_not_true, "System.out.println(!true);", "false");
jt!(logical_not_false, "System.out.println(!false);", "true");
jt!(logical_xor, "System.out.println(true ^ true);", "false");
jt!(
    logical_xor_false,
    "System.out.println(true ^ false);",
    "true"
);
jt!(
    short_circuit_and,
    "System.out.println(false && (10 / 0 == 0));",
    "false"
);
jt!(
    short_circuit_or,
    "System.out.println(true || (10 / 0 == 0));",
    "true"
);
jt!(
    comparison_and_logical,
    "System.out.println(3 > 2 && 5 > 4);",
    "true"
);
jt!(
    comparison_or_logical,
    "System.out.println(3 < 2 || 5 > 4);",
    "true"
);
jt!(
    mixed_precedence,
    "System.out.println(3 > 2 && 2 > 1 || false);",
    "true"
);
jt!(double_negation, "System.out.println(!!true);", "true");
jt!(boolean_from_int_eq, "System.out.println(1 == 1);", "true");
jt!(boolean_from_int_neq, "System.out.println(1 == 2);", "false");
jt!(
    nonnull_reference_comparison,
    "System.out.println(\"a\" != null);",
    "true"
);
jt!(
    null_reference_comparison,
    "String x = null; System.out.println(x == null);",
    "true"
);
jt!(
    object_identity,
    "String a = new String(\"x\"); String b = new String(\"x\"); System.out.println(a == b);",
    "false"
);
jt!(
    object_equality,
    "String a = new String(\"x\"); String b = new String(\"x\"); System.out.println(a.equals(b));",
    "true"
);
jt!(
    ternary_true_branch,
    "System.out.println(true ? \"ok\" : \"no\");",
    "ok"
);
jt!(
    ternary_false_branch,
    "System.out.println(false ? \"ok\" : \"no\");",
    "no"
);
jt!(
    ternary_with_numeric_expr,
    "System.out.println((5 > 3) ? 8 : 2);",
    "8"
);
jt!(
    ternary_chained,
    "System.out.println(true ? (1 > 0 ? 3 : 2) : 0);",
    "3"
);
jt!(
    boolean_chain_many,
    "System.out.println(true && false || false || true);",
    "true"
);
